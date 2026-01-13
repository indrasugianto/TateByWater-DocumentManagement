' Component: **DropboxAPI_POC**
' Type: module
' Lines: 683
' ============================================================

Option Compare Database
Option Explicit

' ============================================================================
' Module: DropboxAPI_POC
' Purpose: Proof of concept for Dropbox API integration
' Date: 2026-01-12
' Your App: jbozj8nffezcw9w
' ============================================================================



' ===== YOUR DROPBOX APP CREDENTIALS =====
Private Const DROPBOX_APP_KEY As String = "jbozj8nffezcw9w"
Private Const DROPBOX_APP_SECRET As String = "qjp2rzxzgfhj9qf"
Private Const DROPBOX_REDIRECT_URI As String = "http://localhost"

' API Endpoints (don't change these)
Private Const API_BASE As String = "https://api.dropboxapi.com/2/"
Private Const CONTENT_BASE As String = "https://content.dropboxapi.com/2/"
Private Const AUTH_URL As String = "https://www.dropbox.com/oauth2/authorize"
Private Const TOKEN_URL As String = "https://api.dropbox.com/oauth2/token"

' Module-level token storage (POC only - will store in database later)
Private m_AccessToken As String
Private m_RefreshToken As String

' ============================================================================
' AUTHENTICATION
' ============================================================================

Public Function AuthenticateUser() As Boolean
    ' Step 1: Build authorization URL
    Dim authURL As String
    authURL = AUTH_URL & "?" & _
              "client_id=" & DROPBOX_APP_KEY & _
              "&response_type=code" & _
              "&token_access_type=offline" & _
              "&redirect_uri=" & DROPBOX_REDIRECT_URI
    
    ' Step 2: Open browser for user authorization
    Debug.Print "Opening authorization URL..."
    Debug.Print authURL
    Application.FollowHyperlink authURL
    
    MsgBox "Your browser will open to Dropbox." & vbCrLf & vbCrLf & _
           "1. Click 'Allow' to authorize the app" & vbCrLf & _
           "2. Copy the authorization code shown" & vbCrLf & _
           "3. Paste it in the next prompt", vbInformation, "Step 1: Authorize"
    
    ' Step 3: Get authorization code from user
    Dim authCode As String
    authCode = InputBox("Paste the authorization code here:", _
                        "Step 2: Authorization Code", "")
    
    If authCode = "" Then
        MsgBox "Authentication cancelled", vbInformation
        AuthenticateUser = False
        Exit Function
    End If
    
    ' Step 4: Exchange code for tokens
    If ExchangeCodeForTokens(authCode) Then
        MsgBox "? Authentication successful!" & vbCrLf & vbCrLf & _
               "You can now use Upload, Download, and other functions.", _
               vbInformation, "Success!"
        AuthenticateUser = True
    Else
        MsgBox "? Authentication failed" & vbCrLf & vbCrLf & _
               "Check the Immediate Window (Ctrl+G) for error details.", _
               vbCritical, "Failed"
        AuthenticateUser = False
    End If
End Function

Private Function ExchangeCodeForTokens(authCode As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim postData As String
    Dim response As String
    
    Debug.Print "Exchanging authorization code for tokens..."
    
    ' Build POST data
    postData = "code=" & authCode & _
               "&grant_type=authorization_code" & _
               "&client_id=" & DROPBOX_APP_KEY & _
               "&client_secret=" & DROPBOX_APP_SECRET & _
               "&redirect_uri=" & DROPBOX_REDIRECT_URI
    
    ' Make token request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", TOKEN_URL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send postData
    
    Debug.Print "Response Status: " & http.Status
    
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "===== FULL TOKEN RESPONSE ====="
        Debug.Print response
        Debug.Print "==============================="
        
        ' Try to extract tokens with better debugging
        m_AccessToken = ExtractJsonValue(response, "access_token")
        m_RefreshToken = ExtractJsonValue(response, "refresh_token")
        
        Debug.Print "Extracted Access Token: [" & m_AccessToken & "]"
        Debug.Print "Extracted Refresh Token: [" & m_RefreshToken & "]"
        
        If m_AccessToken <> "" Then
            Debug.Print "? Access Token: " & Left(m_AccessToken, 20) & "..."
            If m_RefreshToken <> "" Then
                Debug.Print "? Refresh Token: " & Left(m_RefreshToken, 20) & "..."
            End If
            ExchangeCodeForTokens = True
        Else
            Debug.Print "? Failed to extract access_token from response"
            Debug.Print "? Please copy the response above and report it"
            ExchangeCodeForTokens = False
        End If
    Else
        Debug.Print "? Token Error: " & http.Status & " - " & http.responseText
        ExchangeCodeForTokens = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "? Error in ExchangeCodeForTokens: " & Err.Description
    ExchangeCodeForTokens = False
End Function
' ============================================================================
' FILE OPERATIONS
' ============================================================================

Public Function UploadFile(localFilePath As String, dropboxPath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim apiArg As String
    Dim startTime As Double
    Dim fileSize As Long
    
    Debug.Print "========================================="
    Debug.Print "UPLOAD: " & localFilePath
    Debug.Print "TO:     " & dropboxPath
    startTime = Timer
    
    ' Check if token exists
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first (click Authenticate button)", vbExclamation
        UploadFile = False
        Exit Function
    End If
    
    ' Check if file exists
    If Dir(localFilePath) = "" Then
        MsgBox "File not found: " & localFilePath, vbExclamation
        UploadFile = False
        Exit Function
    End If
    
    ' Read file as binary - FIXED METHOD
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' adTypeBinary
    fileStream.Open
    fileStream.LoadFromFile localFilePath
    fileSize = fileStream.Size
    
    Debug.Print "File Size: " & Format(fileSize / 1024, "#,##0") & " KB"
    
    ' Build API argument
    apiArg = "{""path"":""" & dropboxPath & """,""mode"":""overwrite"",""autorename"":false,""mute"":false}"
    
    ' Make API call - FIXED METHOD
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", CONTENT_BASE & "files/upload", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Dropbox-API-Arg", apiArg
    http.setRequestHeader "Content-Type", "application/octet-stream"
    
    ' Send the stream directly (not the Read() result)
    Dim fileData As Variant
    fileData = fileStream.Read
    fileStream.Close
    
    ' Convert to byte array properly
    http.send fileData
    
    Debug.Print "Response Status: " & http.Status
    
    If http.Status = 200 Then
        Debug.Print "? Upload successful! Time: " & Format(Timer - startTime, "0.00") & "s"
        Debug.Print "Response: " & Left(http.responseText, 200) & "..."
        Debug.Print "========================================="
        
        MsgBox "? File uploaded successfully!" & vbCrLf & vbCrLf & _
               "File: " & Dir(localFilePath) & vbCrLf & _
               "Size: " & Format(fileSize / 1024, "#,##0") & " KB" & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & " seconds", _
               vbInformation, "Upload Success"
        UploadFile = True
    Else
        Debug.Print "? Upload failed: " & http.Status
        Debug.Print "Response: " & http.responseText
        Debug.Print "========================================="
        
        MsgBox "? Upload failed!" & vbCrLf & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Error: " & http.responseText, _
               vbCritical, "Upload Failed"
        UploadFile = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "? Error in UploadFile: " & Err.Description
    Debug.Print "? Error Number: " & Err.Number
    Debug.Print "========================================="
    
    ' Close stream if still open
    On Error Resume Next
    If Not fileStream Is Nothing Then
        fileStream.Close
    End If
    
    MsgBox "? Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
    UploadFile = False
End Function
Public Function DownloadFile(dropboxPath As String, localFilePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim apiArg As String
    Dim startTime As Double
    Dim fileSize As Long
    
    Debug.Print "========================================="
    Debug.Print "DOWNLOAD: " & dropboxPath
    Debug.Print "TO:       " & localFilePath
    startTime = Timer
    
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first", vbExclamation
        DownloadFile = False
        Exit Function
    End If
    
    ' Build API argument
    apiArg = "{""path"":""" & dropboxPath & """}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", CONTENT_BASE & "files/download", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Dropbox-API-Arg", apiArg
    
    http.send
    
    Debug.Print "Response Status: " & http.Status
    
    If http.Status = 200 Then
        fileSize = LenB(http.responseBody)
        
        ' Save to file
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Type = 1 ' adTypeBinary
        fileStream.Open
        fileStream.Write http.responseBody
        fileStream.SaveToFile localFilePath, 2 ' adSaveCreateOverWrite
        fileStream.Close
        
        Debug.Print "? Download successful! Time: " & Format(Timer - startTime, "0.00") & "s"
        Debug.Print "File Size: " & Format(fileSize / 1024, "#,##0") & " KB"
        Debug.Print "========================================="
        
        MsgBox "? File downloaded successfully!" & vbCrLf & vbCrLf & _
               "Saved to: " & localFilePath & vbCrLf & _
               "Size: " & Format(fileSize / 1024, "#,##0") & " KB" & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & " seconds", _
               vbInformation, "Download Success"
        DownloadFile = True
    Else
        Debug.Print "? Download failed: " & http.Status
        Debug.Print "Response: " & http.responseText
        Debug.Print "========================================="
        
        MsgBox "? Download failed!" & vbCrLf & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Error: " & http.responseText, _
               vbCritical, "Download Failed"
        DownloadFile = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "? Error in DownloadFile: " & Err.Description
    Debug.Print "========================================="
    MsgBox "? Error: " & Err.Description, vbCritical
    DownloadFile = False
End Function

Public Function CreateFolder(dropboxPath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    
    Debug.Print "========================================="
    Debug.Print "CREATE FOLDER: " & dropboxPath
    
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first", vbExclamation
        CreateFolder = False
        Exit Function
    End If
    
    ' Build JSON body
    jsonBody = "{""path"":""" & dropboxPath & """,""autorename"":false}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", API_BASE & "files/create_folder_v2", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Content-Type", "application/json"
    
    http.send jsonBody
    
    Debug.Print "Response Status: " & http.Status
    
    If http.Status = 200 Then
        Debug.Print "? Folder created successfully"
        Debug.Print "Response: " & http.responseText
        Debug.Print "========================================="
        
        MsgBox "? Folder created: " & dropboxPath, vbInformation, "Success"
        CreateFolder = True
    ElseIf InStr(http.responseText, "path/conflict/folder") > 0 Then
        Debug.Print "? Folder already exists (treating as success)"
        Debug.Print "========================================="
        
        MsgBox "? Folder already exists: " & dropboxPath, vbInformation, "Already Exists"
        CreateFolder = True
    Else
        Debug.Print "? Create folder failed: " & http.Status
        Debug.Print "Response: " & http.responseText
        Debug.Print "========================================="
        
        MsgBox "? Create folder failed!" & vbCrLf & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Check Immediate Window for details", _
               vbCritical, "Failed"
        CreateFolder = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "? Error in CreateFolder: " & Err.Description
    Debug.Print "========================================="
    MsgBox "? Error: " & Err.Description, vbCritical
    CreateFolder = False
End Function

Public Function ListFolder(dropboxPath As String) As String
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    Dim response As String
    
    Debug.Print "========================================="
    Debug.Print "LIST FOLDER: " & dropboxPath
    
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first", vbExclamation
        ListFolder = ""
        Exit Function
    End If
    
    ' Build JSON body (handle empty root)
    If dropboxPath = "" Or dropboxPath = "/" Then
        jsonBody = "{""path"":"""",""recursive"":false,""include_deleted"":false}"
    Else
        jsonBody = "{""path"":""" & dropboxPath & """,""recursive"":false,""include_deleted"":false}"
    End If
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", API_BASE & "files/list_folder", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Content-Type", "application/json"
    
    http.send jsonBody
    
    Debug.Print "Response Status: " & http.Status
    
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "? List folder successful"
        Debug.Print "Response: " & Left(response, 500) & "..."
        Debug.Print "========================================="
        
        ListFolder = response
    Else
        Debug.Print "? List folder failed: " & http.Status
        Debug.Print "Response: " & http.responseText
        Debug.Print "========================================="
        
        MsgBox "? List folder failed!" & vbCrLf & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Check Immediate Window for details", _
               vbCritical, "Failed"
        ListFolder = ""
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "? Error in ListFolder: " & Err.Description
    Debug.Print "========================================="
    MsgBox "? Error: " & Err.Description, vbCritical
    ListFolder = ""
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function ExtractJsonValue(jsonString As String, key As String) As String
    ' Fixed JSON parser - handles spaces after colons (Dropbox uses spaces!)
    Dim startPos As Long
    Dim endPos As Long
    
    ' Pattern 1: "key": "value" (WITH space - Dropbox format)
    Dim pattern1 As String
    pattern1 = """" & key & """: """
    startPos = InStr(jsonString, pattern1)
    
    If startPos > 0 Then
        startPos = startPos + Len(pattern1)
        endPos = InStr(startPos, jsonString, """")
        If endPos > startPos Then
            ExtractJsonValue = Mid(jsonString, startPos, endPos - startPos)
            Exit Function
        End If
    End If
    
    ' Pattern 2: "key":"value" (WITHOUT space - fallback)
    Dim pattern2 As String
    pattern2 = """" & key & """:"""
    startPos = InStr(jsonString, pattern2)
    
    If startPos > 0 Then
        startPos = startPos + Len(pattern2)
        endPos = InStr(startPos, jsonString, """")
        If endPos > startPos Then
            ExtractJsonValue = Mid(jsonString, startPos, endPos - startPos)
            Exit Function
        End If
    End If
    
    ' Not found
    Debug.Print "WARNING: Could not find '" & key & "' in JSON"
    ExtractJsonValue = ""
End Function

Public Function GetAccessToken() As String
    GetAccessToken = m_AccessToken
End Function

Public Function IsAuthenticated() As Boolean
    IsAuthenticated = (m_AccessToken <> "")
End Function

' ============================================================================
' TEST FUNCTIONS - Run these from Immediate Window (Ctrl+G)
' ============================================================================

Public Sub TestAuthentication()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Authentication"
    Debug.Print "========================================="
    
    If AuthenticateUser() Then
        Debug.Print "? TEST PASSED: Authentication successful"
    Else
        Debug.Print "? TEST FAILED: Authentication failed"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub TestCreateFolder()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Create Folder"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "? TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    Dim folderPath As String
    folderPath = "/TB_CMS_POC/TestCase/2023-Smith_John/General"
    
    If CreateFolder(folderPath) Then
        Debug.Print "? TEST PASSED: Folder creation successful"
    Else
        Debug.Print "? TEST FAILED: Folder creation failed"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub TestUpload()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Upload File"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "? TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    ' File picker
    Dim fd As FileDialog
    Dim testFile As String
    Dim dropboxPath As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a file to upload (PDF recommended, < 5MB)"
    fd.AllowMultiSelect = False
    
    If fd.Show = -1 Then
        testFile = fd.SelectedItems(1)
        dropboxPath = "/TB_CMS_POC/" & Dir(testFile)
        
        If UploadFile(testFile, dropboxPath) Then
            Debug.Print "? TEST PASSED: Upload successful"
        Else
            Debug.Print "? TEST FAILED: Upload failed"
        End If
    Else
        Debug.Print "? TEST CANCELLED: No file selected"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub TestListFolder()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: List Folder"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "? TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    Dim folderPath As String
    Dim contents As String
    
    folderPath = "/TB_CMS_POC"
    contents = ListFolder(folderPath)
    
    If contents <> "" Then
        Debug.Print "? TEST PASSED: List folder successful"
        Debug.Print "Contents preview: " & Left(contents, 300) & "..."
    Else
        Debug.Print "? TEST FAILED: List folder failed"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub TestDownload()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Download File"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "? TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    ' First, list folder to see what files exist
    Dim contents As String
    contents = ListFolder("/TB_CMS_POC")
    
    If contents = "" Then
        Debug.Print "? TEST SKIPPED: Folder is empty. Upload a file first with TestUpload."
        Exit Sub
    End If
    
    ' Prompt for file to download
    Dim dropboxPath As String
    Dim localPath As String
    
    dropboxPath = InputBox("Enter the Dropbox path to download:" & vbCrLf & vbCrLf & _
                          "Example: /TB_CMS_POC/yourfile.pdf" & vbCrLf & vbCrLf & _
                          "(Check Dropbox web interface for exact name)", _
                          "Download Test", "/TB_CMS_POC/")
    
    If dropboxPath <> "" Then
        ' Create local path
        localPath = Environ("TEMP") & "\" & Right(dropboxPath, Len(dropboxPath) - InStrRev(dropboxPath, "/"))
        
        If DownloadFile(dropboxPath, localPath) Then
            Debug.Print "? TEST PASSED: Download successful"
            
            ' Try to open the file
            On Error Resume Next
            Application.FollowHyperlink localPath
            Debug.Print "File opened: " & localPath
        Else
            Debug.Print "? TEST FAILED: Download failed"
        End If
    Else
        Debug.Print "? TEST CANCELLED: No path entered"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub RunAllTests()
    Debug.Print ""
    Debug.Print "+-------------------------------------------+"
    Debug.Print "¦  DROPBOX API POC - COMPLETE TEST SUITE  ¦"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    
    If Not IsAuthenticated() Then
        Debug.Print "Starting with authentication..."
        Call TestAuthentication
        
        If Not IsAuthenticated() Then
            Debug.Print ""
            Debug.Print "? CANNOT PROCEED: Authentication failed"
            Debug.Print "  Fix authentication issues before running other tests"
            Exit Sub
        End If
    End If
    
    Debug.Print ""
    Debug.Print "Running test suite..."
    Debug.Print ""
    
    ' Test 1: Create Folder
    Call TestCreateFolder
    
    ' Test 2: Upload
    Call TestUpload
    
    ' Test 3: List Folder
    Call TestListFolder
    
    ' Test 4: Download
    Call TestDownload
    
    Debug.Print ""
    Debug.Print "+-------------------------------------------+"
    Debug.Print "¦         ALL TESTS COMPLETED!              ¦"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    Debug.Print "Review results above. All tests should show ?"
    Debug.Print ""
End Sub

