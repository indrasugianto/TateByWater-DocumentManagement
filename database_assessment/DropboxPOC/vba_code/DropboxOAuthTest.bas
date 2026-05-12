Attribute VB_Name = "DropboxOAuthTest"
Option Compare Database
Option Explicit

' =============================================================================
' DropboxOAuthTest.bas
'
' PURPOSE: Smoke-test the full Dropbox OAuth authorization code flow from
'          MS Access VBA and verify access to the /Company/COMMON team folder.
'
' HOW TO RUN:
'   1. Open this database in MS Access.
'   2. Open the VBA editor (Alt+F11).
'   3. Import this file as a module (File > Import File).
'   4. Fill in APP_SECRET below with the current app secret.
'   5. Ensure http://localhost:8765 is registered in the Dropbox App Console
'      (Settings > OAuth 2 > Redirect URIs).
'   6. Run TestDropboxOAuth from the Immediate Window or press F5.
'
' OAUTH APPROACH: local HTTP listener (no copy-paste required)
'   VBA shells a PowerShell HttpListener on localhost:8765 before opening the
'   browser. When the user clicks Allow, Dropbox redirects to localhost:8765
'   automatically. PowerShell captures the code, shows a success page in the
'   browser, and writes the redirect URL to a temp file. VBA polls for that
'   file, then extracts code and state without any user interaction.
'
'   Fallback: if USE_LOCAL_LISTENER = False, reverts to the manual paste flow
'   (redirect_uri = http://localhost, no port, user pastes URL into InputBox).
'
' NOTE: This module uses hardcoded constants for App key/secret.
'       Production DropboxService.bas reads these from tblDropboxConfig
'       (DPAPI-encrypted). Do NOT use this credential pattern in production.
' =============================================================================

' -----------------------------------------------------------------------------
' CONFIGURATION — fill in APP_SECRET before running
' -----------------------------------------------------------------------------
Private Const APP_KEY             As String = "dqleswbnux8k3m5"
Private Const APP_SECRET          As String = "PASTE_YOUR_APP_SECRET_HERE"
Private Const LISTENER_PORT       As Long = 8765
Private Const LISTENER_TIMEOUT_S  As Long = 120     ' seconds to wait for browser redirect
Private Const TEAM_NAMESPACE_ID   As String = "14334595683"
Private Const TEST_FOLDER_PATH    As String = "/Company/COMMON"

' Set to False to fall back to the manual URL-paste flow
Private Const USE_LOCAL_LISTENER  As Boolean = True

Private Const TOKEN_ENDPOINT  As String = "https://api.dropbox.com/oauth2/token"
Private Const AUTH_ENDPOINT   As String = "https://www.dropbox.com/oauth2/authorize"
Private Const API_BASE        As String = "https://api.dropboxapi.com/2"

Private m_OAuthState As String


' =============================================================================
' MAIN ENTRY POINT
' =============================================================================
Public Sub TestDropboxOAuth()
    If APP_SECRET = "PASTE_YOUR_APP_SECRET_HERE" Then
        MsgBox "Fill in APP_SECRET in the DropboxOAuthTest module before running.", _
               vbCritical, "Configuration Required"
        Exit Sub
    End If

    MsgBox "Dropbox OAuth Smoke Test" & vbCrLf & vbCrLf & _
           "Your browser will open to the Dropbox authorization page." & vbCrLf & _
           "Sign in with your Tate Bywater Dropbox account and click Allow." & vbCrLf & vbCrLf & _
           IIf(USE_LOCAL_LISTENER, _
               "The authorization will complete automatically — no copy-paste needed.", _
               "After clicking Allow, copy the full redirect URL and paste it when prompted.") & _
           vbCrLf & vbCrLf & "Click OK to open the browser.", _
           vbInformation, "Dropbox OAuth Test"

    ' Step 1: Get authorization code (listener or paste)
    Dim authCode As String
    If USE_LOCAL_LISTENER Then
        authCode = GetAuthorizationCodeViaListener()
    Else
        authCode = GetAuthorizationCodeViaPaste()
    End If

    If authCode = "" Then
        MsgBox "Authorization was cancelled or failed. Test aborted.", _
               vbExclamation, "OAuth Test"
        Exit Sub
    End If

    ' Step 2: Exchange code for tokens
    Dim tokenJson As String
    tokenJson = ExchangeCodeForToken(authCode)
    If tokenJson = "" Then Exit Sub

    Dim accessToken As String
    Dim refreshToken As String
    accessToken = ExtractJsonString(tokenJson, "access_token")
    refreshToken = ExtractJsonString(tokenJson, "refresh_token")

    If accessToken = "" Then
        MsgBox "Token exchange succeeded but access_token was not found in response." & _
               vbCrLf & vbCrLf & "Response: " & Left(tokenJson, 500), _
               vbCritical, "Token Parse Error"
        Exit Sub
    End If

    ' Step 3: Identity check
    Dim accountEmail As String
    accountEmail = GetAccountEmail(accessToken)
    If accountEmail = "" Then Exit Sub

    ' Step 4: Folder access with namespace header
    Dim folderResult As String
    folderResult = ListFolder(accessToken, TEST_FOLDER_PATH)
    If folderResult = "" Then Exit Sub

    Dim folderCount As Long
    Dim fileCount As Long
    folderCount = CountJsonTag(folderResult, "folder")
    fileCount = CountJsonTag(folderResult, "file")

    MsgBox "ALL TESTS PASSED" & vbCrLf & vbCrLf & _
           "Account:              " & accountEmail & vbCrLf & _
           "Namespace ID:         " & TEAM_NAMESPACE_ID & vbCrLf & _
           "Folder tested:        " & TEST_FOLDER_PATH & vbCrLf & _
           "Subfolders found:     " & folderCount & vbCrLf & _
           "Files found:          " & fileCount & vbCrLf & _
           "Refresh token:        " & IIf(refreshToken <> "", "Obtained", "Missing") & vbCrLf & _
           "Auth method:          " & IIf(USE_LOCAL_LISTENER, "Local HTTP listener", "Manual paste") & _
           vbCrLf & vbCrLf & _
           "OAuth flow and team namespace access are working correctly." & vbCrLf & _
           "Safe to proceed with Phase 3 DropboxService.bas development.", _
           vbInformation, "Dropbox OAuth Test - PASSED"
End Sub


' =============================================================================
' APPROACH A: Local HTTP listener — no user copy-paste required
'
' How it works:
'   1. Write a PowerShell script to a temp .ps1 file
'   2. Shell PowerShell in hidden mode to run that script
'      PowerShell starts an HttpListener on localhost:LISTENER_PORT,
'      waits for one inbound request (the Dropbox redirect), sends a
'      success HTML page to the browser, writes the full redirect URL
'      to a second temp file, then exits.
'   3. Open the browser to the Dropbox auth URL
'   4. Poll the output temp file until it has content or timeout expires
'   5. Read the redirect URL, validate state, extract code
' =============================================================================
Private Function GetAuthorizationCodeViaListener() As String
    ' Generate GUID state parameter
    Dim stateValue As String
    stateValue = GenerateState()
    m_OAuthState = stateValue

    ' Temp file paths
    Dim tempDir As String
    tempDir = Environ("TEMP") & "\TBCMS"
    CreateDirIfNotExists tempDir

    Dim ps1Path As String
    Dim outputPath As String
    ps1Path = tempDir & "\oauth_listener.ps1"
    outputPath = tempDir & "\oauth_result.txt"

    ' Remove stale result file from any previous run
    If Dir(outputPath) <> "" Then Kill outputPath

    ' Write the PowerShell listener script
    If Not WriteListenerScript(ps1Path, outputPath, LISTENER_PORT) Then
        MsgBox "Failed to write PowerShell listener script to:" & vbCrLf & ps1Path, _
               vbCritical, "Listener Setup Error"
        GetAuthorizationCodeViaListener = ""
        Exit Function
    End If

    ' Shell PowerShell listener in background (hidden window)
    Dim psCmd As String
    psCmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
            ps1Path & """ """ & outputPath & """"
    Shell psCmd, vbHide

    ' Brief pause to let the listener start before opening browser
    WaitMilliseconds 800

    ' Build auth URL using the port-specific redirect URI
    Dim redirectUri As String
    redirectUri = "http://localhost:" & LISTENER_PORT

    Dim authUrl As String
    authUrl = AUTH_ENDPOINT & _
              "?client_id=" & APP_KEY & _
              "&response_type=code" & _
              "&token_access_type=offline" & _
              "&state=" & stateValue & _
              "&redirect_uri=" & UrlEncode(redirectUri)

    Application.FollowHyperlink authUrl

    ' Poll for the output file
    Dim redirectUrl As String
    redirectUrl = PollForFile(outputPath, LISTENER_TIMEOUT_S)

    ' Clean up temp files
    On Error Resume Next
    Kill ps1Path
    Kill outputPath
    On Error GoTo 0

    If redirectUrl = "" Then
        MsgBox "Timed out waiting for Dropbox to redirect." & vbCrLf & vbCrLf & _
               "Possible causes:" & vbCrLf & _
               "  - http://localhost:" & LISTENER_PORT & " is not registered in the Dropbox App Console" & vbCrLf & _
               "  - The browser authorization was not completed within " & LISTENER_TIMEOUT_S & " seconds" & vbCrLf & _
               "  - Port " & LISTENER_PORT & " is blocked by a firewall or already in use" & vbCrLf & vbCrLf & _
               "Try setting USE_LOCAL_LISTENER = False to use the manual paste flow instead.", _
               vbExclamation, "Listener Timeout"
        GetAuthorizationCodeViaListener = ""
        Exit Function
    End If

    ' Validate state and extract code
    GetAuthorizationCodeViaListener = ValidateAndExtractCode(redirectUrl)
End Function


' Writes the PowerShell HttpListener script to disk
Private Function WriteListenerScript(ps1Path As String, outputPath As String, _
                                     port As Long) As Boolean
    Dim fileNum As Integer
    fileNum = FreeFile

    On Error GoTo WriteError

    Open ps1Path For Output As #fileNum
    Print #fileNum, "param([string]$OutputFile)"
    Print #fileNum, "$listener = New-Object System.Net.HttpListener"
    Print #fileNum, "$listener.Prefixes.Add('http://localhost:" & port & "/')"
    Print #fileNum, "try {"
    Print #fileNum, "    $listener.Start()"
    Print #fileNum, "    $context = $listener.GetContext()"
    Print #fileNum, "    $redirectUrl = $context.Request.Url.ToString()"
    Print #fileNum, "    $html = '<html><head><style>body{font-family:sans-serif;"
    Print #fileNum, "text-align:center;padding:60px;background:#f0f4f8}"
    Print #fileNum, "h2{color:#2d7d46}p{color:#444}</style></head><body>"
    Print #fileNum, "<h2>&#10003; Authorization Complete</h2>"
    Print #fileNum, "<p>You can close this tab and return to TBCMS in Access.</p>"
    Print #fileNum, "</body></html>'"
    Print #fileNum, "    $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)"
    Print #fileNum, "    $context.Response.ContentType = 'text/html; charset=utf-8'"
    Print #fileNum, "    $context.Response.ContentLength64 = $bytes.Length"
    Print #fileNum, "    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)"
    Print #fileNum, "    $context.Response.OutputStream.Close()"
    Print #fileNum, "    $redirectUrl | Out-File -FilePath $OutputFile -Encoding UTF8 -NoNewline"
    Print #fileNum, "} finally {"
    Print #fileNum, "    $listener.Stop()"
    Print #fileNum, "}"
    Close #fileNum

    WriteListenerScript = True
    Exit Function

WriteError:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    WriteListenerScript = False
End Function


' Poll a file path until it has content, or timeout expires
' Returns file contents (trimmed), or "" on timeout
Private Function PollForFile(filePath As String, timeoutSeconds As Long) As String
    Dim startTime As Single
    startTime = Timer

    Dim content As String
    Dim fileNum As Integer

    Do
        If Dir(filePath) <> "" Then
            ' File exists — try to read it
            On Error Resume Next
            fileNum = FreeFile
            Open filePath For Input As #fileNum
            content = ""
            Dim line As String
            Do While Not EOF(fileNum)
                Line Input #fileNum, line
                content = content & line
            Loop
            Close #fileNum
            On Error GoTo 0

            content = Trim(content)
            If content <> "" Then
                PollForFile = content
                Exit Function
            End If
        End If

        WaitMilliseconds 500

        ' Handle Timer midnight rollover
        Dim elapsed As Single
        elapsed = Timer - startTime
        If elapsed < 0 Then elapsed = elapsed + 86400

    Loop While elapsed < timeoutSeconds

    PollForFile = ""
End Function


' =============================================================================
' APPROACH B: Manual paste fallback
' =============================================================================
Private Function GetAuthorizationCodeViaPaste() As String
    Dim stateValue As String
    stateValue = GenerateState()
    m_OAuthState = stateValue

    Dim authUrl As String
    authUrl = AUTH_ENDPOINT & _
              "?client_id=" & APP_KEY & _
              "&response_type=code" & _
              "&token_access_type=offline" & _
              "&state=" & stateValue & _
              "&redirect_uri=" & UrlEncode("http://localhost")

    Application.FollowHyperlink authUrl

    Dim redirectUrl As String
    redirectUrl = InputBox( _
        "Your browser has opened the Dropbox authorization page." & vbCrLf & vbCrLf & _
        "Steps:" & vbCrLf & _
        "  1. Sign in with your Tate Bywater Dropbox account" & vbCrLf & _
        "  2. Click Allow" & vbCrLf & _
        "  3. The browser will show an error page — that is expected" & vbCrLf & _
        "  4. Copy the FULL URL from the browser address bar and paste it below:", _
        "Paste Redirect URL")

    If Trim(redirectUrl) = "" Then
        GetAuthorizationCodeViaPaste = ""
        Exit Function
    End If

    GetAuthorizationCodeViaPaste = ValidateAndExtractCode(redirectUrl)
End Function


' =============================================================================
' SHARED: validate state and extract code from redirect URL
' =============================================================================
Private Function ValidateAndExtractCode(redirectUrl As String) As String
    Dim errorParam As String
    errorParam = ExtractQueryParam(redirectUrl, "error")
    If errorParam <> "" Then
        MsgBox "Dropbox returned an error: " & errorParam & vbCrLf & _
               ExtractQueryParam(redirectUrl, "error_description"), _
               vbCritical, "Authorization Error"
        ValidateAndExtractCode = ""
        Exit Function
    End If

    Dim returnedState As String
    returnedState = ExtractQueryParam(redirectUrl, "state")
    If returnedState <> m_OAuthState Then
        MsgBox "State parameter mismatch — possible CSRF. Authorization aborted." & _
               vbCrLf & vbCrLf & _
               "Expected: " & m_OAuthState & vbCrLf & _
               "Received: " & returnedState, _
               vbCritical, "Security Error"
        ValidateAndExtractCode = ""
        Exit Function
    End If

    Dim code As String
    code = ExtractQueryParam(redirectUrl, "code")
    If code = "" Then
        MsgBox "Could not extract authorization code from redirect URL." & vbCrLf & _
               "URL received: " & Left(redirectUrl, 300), _
               vbCritical, "Parse Error"
    End If

    ValidateAndExtractCode = code
End Function


' =============================================================================
' STEP 2: Exchange authorization code for access + refresh tokens
' =============================================================================
Private Function ExchangeCodeForToken(authCode As String) As String
    Dim redirectUri As String
    If USE_LOCAL_LISTENER Then
        redirectUri = "http://localhost:" & LISTENER_PORT
    Else
        redirectUri = "http://localhost"
    End If

    Dim postBody As String
    postBody = "code=" & authCode & _
               "&grant_type=authorization_code" & _
               "&client_id=" & APP_KEY & _
               "&client_secret=" & APP_SECRET & _
               "&redirect_uri=" & UrlEncode(redirectUri)

    Dim responseText As String
    responseText = HttpPost(TOKEN_ENDPOINT, postBody, "application/x-www-form-urlencoded", "")

    If responseText = "" Then
        ExchangeCodeForToken = ""
        Exit Function
    End If

    If InStr(responseText, "access_token") = 0 Then
        MsgBox "Token exchange failed." & vbCrLf & vbCrLf & _
               "Response: " & Left(responseText, 500), _
               vbCritical, "Token Exchange Error"
        ExchangeCodeForToken = ""
        Exit Function
    End If

    ExchangeCodeForToken = responseText
End Function


' =============================================================================
' STEP 3: Get account email to verify correct identity
' =============================================================================
Private Function GetAccountEmail(accessToken As String) As String
    ' Note: namespace header is NOT required for users/get_current_account
    Dim responseText As String
    responseText = HttpPost(API_BASE & "/users/get_current_account", "null", _
                            "application/json", accessToken)
    If responseText = "" Then
        GetAccountEmail = ""
        Exit Function
    End If

    Dim email As String
    email = ExtractJsonString(responseText, "email")
    If email = "" Then
        MsgBox "Could not extract email from account response." & vbCrLf & _
               "Response: " & Left(responseText, 300), _
               vbCritical, "Account Check Error"
    End If

    GetAccountEmail = email
End Function


' =============================================================================
' STEP 4: List a folder using the team namespace header
' =============================================================================
Private Function ListFolder(accessToken As String, folderPath As String) As String
    Dim requestBody As String
    requestBody = "{""path"": """ & folderPath & """, ""recursive"": false}"

    Dim namespaceHeader As String
    namespaceHeader = "{""namespace_id"": """ & TEAM_NAMESPACE_ID & _
                      """, "".tag"": ""namespace_id""}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo HttpError

    http.Open "POST", API_BASE & "/files/list_folder", False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Dropbox-API-Path-Root", namespaceHeader
    http.Send requestBody

    If http.Status <> 200 Then
        MsgBox "Folder listing failed." & vbCrLf & _
               "HTTP " & http.Status & vbCrLf & vbCrLf & _
               "Response: " & Left(http.ResponseText, 500), _
               vbCritical, "Folder Access Error"
        ListFolder = ""
        Exit Function
    End If

    ListFolder = http.ResponseText
    Exit Function

HttpError:
    MsgBox "HTTP error listing folder: " & Err.Description, vbCritical, "Folder Access Error"
    ListFolder = ""
End Function


' =============================================================================
' HTTP HELPER
' =============================================================================
Private Function HttpPost(url As String, body As String, _
                          contentType As String, bearerToken As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo HttpError

    http.Open "POST", url, False
    If bearerToken <> "" Then
        http.SetRequestHeader "Authorization", "Bearer " & bearerToken
    End If
    http.SetRequestHeader "Content-Type", contentType
    http.Send body

    If http.Status < 200 Or http.Status >= 300 Then
        MsgBox "HTTP request to " & url & " failed." & vbCrLf & _
               "HTTP " & http.Status & vbCrLf & vbCrLf & _
               Left(http.ResponseText, 500), _
               vbCritical, "HTTP Error"
        HttpPost = ""
        Exit Function
    End If

    HttpPost = http.ResponseText
    Exit Function

HttpError:
    MsgBox "Network error calling " & url & ":" & vbCrLf & Err.Description, _
           vbCritical, "HTTP Error"
    HttpPost = ""
End Function


' =============================================================================
' UTILITY HELPERS
' =============================================================================

Private Function GenerateState() As String
    Dim guid As String
    On Error Resume Next
    guid = CreateObject("Scriptlet.TypeLib").guid
    On Error GoTo 0
    If guid = "" Then
        guid = "tbcms_" & Format(Now(), "yyyymmddhhmmss") & "_" & Int(Rnd() * 99999)
    End If
    GenerateState = Replace(Replace(guid, "{", ""), "}", "")
End Function

Private Sub CreateDirIfNotExists(dirPath As String)
    If Dir(dirPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir dirPath
        On Error GoTo 0
    End If
End Sub

' Non-blocking sleep using a hidden WScript.Shell trick via DoEvents loop
Private Sub WaitMilliseconds(ms As Long)
    Dim endTime As Single
    endTime = Timer + (ms / 1000)
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

Private Function ExtractQueryParam(url As String, paramName As String) As String
    Dim queryString As String
    Dim questionPos As Long
    questionPos = InStr(url, "?")
    If questionPos = 0 Then
        ExtractQueryParam = ""
        Exit Function
    End If

    queryString = Mid(url, questionPos + 1)
    Dim hashPos As Long
    hashPos = InStr(queryString, "#")
    If hashPos > 0 Then queryString = Left(queryString, hashPos - 1)

    Dim parts() As String
    parts = Split(queryString, "&")
    Dim i As Integer
    Dim prefix As String
    prefix = paramName & "="
    For i = 0 To UBound(parts)
        If Left(parts(i), Len(prefix)) = prefix Then
            ExtractQueryParam = Mid(parts(i), Len(prefix) + 1)
            Exit Function
        End If
    Next i
    ExtractQueryParam = ""
End Function

Private Function ExtractJsonString(json As String, key As String) As String
    Dim searchKey As String
    searchKey = """" & key & """:"
    Dim startPos As Long
    startPos = InStr(json, searchKey)
    If startPos = 0 Then
        ExtractJsonString = ""
        Exit Function
    End If
    startPos = startPos + Len(searchKey)
    Do While startPos <= Len(json) And _
             (Mid(json, startPos, 1) = " " Or Mid(json, startPos, 1) = Chr(9))
        startPos = startPos + 1
    Loop
    If startPos > Len(json) Or Mid(json, startPos, 1) <> """" Then
        ExtractJsonString = ""
        Exit Function
    End If
    startPos = startPos + 1
    Dim endPos As Long
    endPos = startPos
    Do While endPos <= Len(json)
        If Mid(json, endPos, 1) = "\" Then
            endPos = endPos + 2
        ElseIf Mid(json, endPos, 1) = """" Then
            Exit Do
        Else
            endPos = endPos + 1
        End If
    Loop
    ExtractJsonString = Mid(json, startPos, endPos - startPos)
End Function

Private Function CountJsonTag(json As String, tagValue As String) As Long
    Dim searchStr As String
    searchStr = """" & tagValue & """"
    Dim count As Long
    Dim pos As Long
    count = 0
    pos = 1
    Do
        pos = InStr(pos, json, searchStr)
        If pos = 0 Then Exit Do
        count = count + 1
        pos = pos + Len(searchStr)
    Loop
    CountJsonTag = count
End Function

Private Function UrlEncode(text As String) As String
    Dim result As String
    Dim i As Integer
    Dim charCode As Integer
    Dim c As String
    result = ""
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        charCode = Asc(c)
        Select Case charCode
            Case 65 To 90, 97 To 122, 48 To 57
                result = result & c
            Case 45, 46, 95, 126
                result = result & c
            Case Else
                result = result & "%" & Right("0" & Hex(charCode), 2)
        End Select
    Next i
    UrlEncode = result
End Function
