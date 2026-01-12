# Dropbox API Proof of Concept Plan

**Purpose**: Validate core Dropbox API integration approach before full development  
**Timeline**: 1-2 weeks  
**Goal**: Working prototype with essential upload/download functionality  
**Date**: 2026-01-12

---

## POC Objectives

### Primary Goals
1. ✅ **Validate OAuth 2.0** - Prove authentication works in VBA
2. ✅ **Test File Upload** - Successfully upload documents to Dropbox
3. ✅ **Test File Download** - Retrieve and open documents
4. ✅ **Test Folder Creation** - Create case folder structure
5. ✅ **Measure Performance** - Understand API response times
6. ✅ **Identify Challenges** - Surface technical issues early

### Success Criteria
- ✅ Successfully authenticate with Dropbox
- ✅ Upload a document (< 5 seconds for 1MB file)
- ✅ Download and open a document
- ✅ Create nested folder structure
- ✅ Handle basic errors gracefully
- ✅ Document lessons learned

---

## Scope

### ✅ IN SCOPE (Must Have)

**Core Functions:**
1. OAuth 2.0 authentication
2. Upload single file
3. Download single file
4. Create folder
5. List folder contents
6. Basic error handling

**Test Scenarios:**
- Upload document from scanner folder
- Download and open document
- Create case folder structure
- Handle authentication expiry
- Handle network error

### ❌ OUT OF SCOPE (Later)

**Deferred to Full Implementation:**
- Full DocumentManagement.bas integration
- Local caching system
- Chunked uploads (large files)
- Version history
- Document sharing
- Advanced error handling
- Retry logic with exponential backoff
- Rate limiting management
- Database integration
- UI modifications
- Migration scripts
- Operation queueing

---

## Architecture (POC Version)

```
┌─────────────────────────────────────────┐
│      Simple Test Form (NEW)             │
│  ┌────────────────────────────────┐    │
│  │  [Authenticate]                 │    │
│  │  [Upload File]                  │    │
│  │  [Download File]                │    │
│  │  [Create Folder]                │    │
│  │  [List Folder]                  │    │
│  │                                 │    │
│  │  Status: _____________          │    │
│  └────────────────────────────────┘    │
│                │                        │
│                ▼                        │
│  ┌────────────────────────────────┐    │
│  │   DropboxAPI_POC.bas           │────┼──→ Dropbox API
│  │   (Simplified API module)      │    │    (HTTPS REST)
│  └────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

**Simplified Approach:**
- Single test form for manual testing
- One simplified API module
- No database integration (test with hardcoded paths)
- No caching (direct API calls)
- Minimal error handling (just show errors)

---

## Implementation Plan

### Phase 1: Setup (Day 1)

#### Step 1.1: Dropbox App Registration (2 hours)
```
Tasks:
□ Go to https://www.dropbox.com/developers/apps
□ Click "Create app"
□ Choose "Scoped access"
□ Choose "Full Dropbox" access
□ Name: "TB CMS POC"
□ Click "Create app"
□ In app settings:
  □ Copy App Key
  □ Copy App Secret
  □ Under "OAuth 2" → Add redirect URI: http://localhost
  □ Under "Permissions" → Grant these scopes:
    - files.metadata.write
    - files.metadata.read
    - files.content.write
    - files.content.read
  □ Click "Submit"
□ Document credentials securely

Deliverable: App Key, App Secret, redirect URI
```

#### Step 1.2: Test Dropbox Account (1 hour)
```
Tasks:
□ Create test folder in Dropbox: "/TB_CMS_POC"
□ Create subfolder: "/TB_CMS_POC/TestCase"
□ Upload a test document manually
□ Verify folder structure
□ Test from web interface

Deliverable: Test Dropbox structure ready
```

#### Step 1.3: Development Environment (1 hour)
```
Tasks:
□ Open MS Access
□ Create new blank database: "DropboxPOC.accdb"
□ Enable Trust Center settings for VBA
□ Add reference: Microsoft XML, v6.0 (MSXML2)
□ Add reference: Microsoft Scripting Runtime
□ Create module: DropboxAPI_POC
□ Create form: frmDropboxPOC

Deliverable: Development environment ready
```

---

### Phase 2: Core API Module (Days 2-3)

#### Step 2.1: Create DropboxAPI_POC Module

**File**: `DropboxAPI_POC.bas`

```vba
' ============================================================================
' Module: DropboxAPI_POC
' Purpose: Proof of concept for Dropbox API integration
' Date: 2026-01-12
' ============================================================================

Option Compare Database
Option Explicit

' API Configuration - REPLACE WITH YOUR VALUES
Private Const DROPBOX_APP_KEY As String = "YOUR_APP_KEY_HERE"
Private Const DROPBOX_APP_SECRET As String = "YOUR_APP_SECRET_HERE"
Private Const DROPBOX_REDIRECT_URI As String = "http://localhost"

' API Endpoints
Private Const API_BASE As String = "https://api.dropboxapi.com/2/"
Private Const CONTENT_BASE As String = "https://content.dropboxapi.com/2/"
Private Const AUTH_URL As String = "https://www.dropbox.com/oauth2/authorize"
Private Const TOKEN_URL As String = "https://api.dropbox.com/oauth2/token"

' Module-level token storage (POC only - store in database later)
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
    Application.FollowHyperlink authURL
    
    ' Step 3: Get authorization code from user
    Dim authCode As String
    authCode = InputBox("After authorizing, paste the authorization code here:", _
                        "Dropbox Authorization", "")
    
    If authCode = "" Then
        MsgBox "Authentication cancelled", vbInformation
        AuthenticateUser = False
        Exit Function
    End If
    
    ' Step 4: Exchange code for tokens
    If ExchangeCodeForTokens(authCode) Then
        MsgBox "Authentication successful!" & vbCrLf & _
               "Access Token: " & Left(m_AccessToken, 20) & "...", _
               vbInformation, "Success"
        AuthenticateUser = True
    Else
        MsgBox "Authentication failed", vbCritical
        AuthenticateUser = False
    End If
End Function

Private Function ExchangeCodeForTokens(authCode As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim postData As String
    Dim response As String
    
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
    
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "Token Response: " & response
        
        ' Extract tokens (simple string parsing for POC)
        m_AccessToken = ExtractJsonValue(response, "access_token")
        m_RefreshToken = ExtractJsonValue(response, "refresh_token")
        
        If m_AccessToken <> "" Then
            ExchangeCodeForTokens = True
        Else
            ExchangeCodeForTokens = False
        End If
    Else
        Debug.Print "Token Error: " & http.Status & " - " & http.responseText
        ExchangeCodeForTokens = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "Error in ExchangeCodeForTokens: " & Err.Description
    ExchangeCodeForTokens = False
End Function

' ============================================================================
' FILE OPERATIONS
' ============================================================================

Public Function UploadFile(localFilePath As String, dropboxPath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim fileBytes() As Byte
    Dim apiArg As String
    Dim startTime As Double
    
    Debug.Print "Uploading: " & localFilePath & " to " & dropboxPath
    startTime = Timer
    
    ' Check if token exists
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first", vbExclamation
        UploadFile = False
        Exit Function
    End If
    
    ' Read file as binary
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' adTypeBinary
    fileStream.Open
    fileStream.LoadFromFile localFilePath
    fileBytes = fileStream.Read
    fileStream.Close
    
    ' Build API argument
    apiArg = "{""path"":""" & dropboxPath & """,""mode"":""overwrite"",""autorename"":false,""mute"":false}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", CONTENT_BASE & "files/upload", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Dropbox-API-Arg", apiArg
    http.setRequestHeader "Content-Type", "application/octet-stream"
    
    http.send fileBytes
    
    If http.Status = 200 Then
        Debug.Print "Upload successful! Time: " & Format(Timer - startTime, "0.00") & "s"
        Debug.Print "Response: " & http.responseText
        MsgBox "File uploaded successfully!" & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & " seconds", _
               vbInformation
        UploadFile = True
    Else
        Debug.Print "Upload failed: " & http.Status & " - " & http.responseText
        MsgBox "Upload failed: " & http.Status & vbCrLf & http.responseText, vbCritical
        UploadFile = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "Error in UploadFile: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
    UploadFile = False
End Function

Public Function DownloadFile(dropboxPath As String, localFilePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim apiArg As String
    Dim startTime As Double
    
    Debug.Print "Downloading: " & dropboxPath & " to " & localFilePath
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
    
    If http.Status = 200 Then
        ' Save to file
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Type = 1 ' adTypeBinary
        fileStream.Open
        fileStream.Write http.responseBody
        fileStream.SaveToFile localFilePath, 2 ' adSaveCreateOverWrite
        fileStream.Close
        
        Debug.Print "Download successful! Time: " & Format(Timer - startTime, "0.00") & "s"
        MsgBox "File downloaded successfully!" & vbCrLf & _
               "Saved to: " & localFilePath & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & " seconds", _
               vbInformation
        DownloadFile = True
    Else
        Debug.Print "Download failed: " & http.Status & " - " & http.responseText
        MsgBox "Download failed: " & http.Status & vbCrLf & http.responseText, vbCritical
        DownloadFile = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "Error in DownloadFile: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
    DownloadFile = False
End Function

Public Function CreateFolder(dropboxPath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    
    Debug.Print "Creating folder: " & dropboxPath
    
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
    
    If http.Status = 200 Then
        Debug.Print "Folder created successfully"
        Debug.Print "Response: " & http.responseText
        MsgBox "Folder created: " & dropboxPath, vbInformation
        CreateFolder = True
    ElseIf InStr(http.responseText, "path/conflict/folder") > 0 Then
        Debug.Print "Folder already exists - treating as success"
        MsgBox "Folder already exists: " & dropboxPath, vbInformation
        CreateFolder = True
    Else
        Debug.Print "Create folder failed: " & http.Status & " - " & http.responseText
        MsgBox "Create folder failed: " & http.Status & vbCrLf & http.responseText, vbCritical
        CreateFolder = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "Error in CreateFolder: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
    CreateFolder = False
End Function

Public Function ListFolder(dropboxPath As String) As String
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    Dim response As String
    
    Debug.Print "Listing folder: " & dropboxPath
    
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first", vbExclamation
        ListFolder = ""
        Exit Function
    End If
    
    ' Build JSON body
    jsonBody = "{""path"":""" & dropboxPath & """,""recursive"":false,""include_deleted"":false}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", API_BASE & "files/list_folder", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Content-Type", "application/json"
    
    http.send jsonBody
    
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "List folder successful"
        Debug.Print "Response: " & Left(response, 500) & "..."
        ListFolder = response
    Else
        Debug.Print "List folder failed: " & http.Status & " - " & http.responseText
        MsgBox "List folder failed: " & http.Status & vbCrLf & http.responseText, vbCritical
        ListFolder = ""
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "Error in ListFolder: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
    ListFolder = ""
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function ExtractJsonValue(jsonString As String, key As String) As String
    ' Simple JSON value extraction for POC
    ' Find: "key":"value"
    Dim startPos As Long
    Dim endPos As Long
    Dim searchStr As String
    
    searchStr = """" & key & """:"""
    startPos = InStr(jsonString, searchStr)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        endPos = InStr(startPos, jsonString, """")
        If endPos > startPos Then
            ExtractJsonValue = Mid(jsonString, startPos, endPos - startPos)
        End If
    End If
End Function

Public Function GetAccessToken() As String
    GetAccessToken = m_AccessToken
End Function

Public Function IsAuthenticated() As Boolean
    IsAuthenticated = (m_AccessToken <> "")
End Function

' ============================================================================
' TEST FUNCTIONS
' ============================================================================

Public Sub TestAuthentication()
    Debug.Print "=== Test Authentication ==="
    If AuthenticateUser() Then
        Debug.Print "✓ Authentication successful"
    Else
        Debug.Print "✗ Authentication failed"
    End If
End Sub

Public Sub TestUpload()
    Debug.Print "=== Test Upload ==="
    Dim testFile As String
    Dim dropboxPath As String
    
    ' Prompt for file
    testFile = Application.FileDialog(msoFileDialogFilePicker).Show
    If testFile = "" Then
        Debug.Print "No file selected"
        Exit Sub
    End If
    
    dropboxPath = "/TB_CMS_POC/test_upload.pdf"
    
    If UploadFile(testFile, dropboxPath) Then
        Debug.Print "✓ Upload successful"
    Else
        Debug.Print "✗ Upload failed"
    End If
End Sub

Public Sub TestDownload()
    Debug.Print "=== Test Download ==="
    Dim dropboxPath As String
    Dim localPath As String
    
    dropboxPath = "/TB_CMS_POC/test_upload.pdf"
    localPath = "C:\Temp\downloaded_test.pdf"
    
    If DownloadFile(dropboxPath, localPath) Then
        Debug.Print "✓ Download successful"
        ' Open downloaded file
        Application.FollowHyperlink localPath
    Else
        Debug.Print "✗ Download failed"
    End If
End Sub

Public Sub TestCreateFolder()
    Debug.Print "=== Test Create Folder ==="
    Dim folderPath As String
    
    folderPath = "/TB_CMS_POC/TestCase/2023-Smith_John/General"
    
    If CreateFolder(folderPath) Then
        Debug.Print "✓ Folder creation successful"
    Else
        Debug.Print "✗ Folder creation failed"
    End If
End Sub

Public Sub TestListFolder()
    Debug.Print "=== Test List Folder ==="
    Dim folderPath As String
    Dim contents As String
    
    folderPath = "/TB_CMS_POC"
    contents = ListFolder(folderPath)
    
    If contents <> "" Then
        Debug.Print "✓ List folder successful"
        Debug.Print "Contents: " & contents
    Else
        Debug.Print "✗ List folder failed"
    End If
End Sub

Public Sub RunAllTests()
    Debug.Print "======================================="
    Debug.Print "    DROPBOX API POC - TEST SUITE"
    Debug.Print "======================================="
    Debug.Print ""
    
    If Not IsAuthenticated() Then
        TestAuthentication
        If Not IsAuthenticated() Then
            Debug.Print "Cannot proceed without authentication"
            Exit Sub
        End If
    End If
    
    Debug.Print ""
    TestCreateFolder
    
    Debug.Print ""
    TestUpload
    
    Debug.Print ""
    TestListFolder
    
    Debug.Print ""
    TestDownload
    
    Debug.Print ""
    Debug.Print "======================================="
    Debug.Print "    TESTS COMPLETE"
    Debug.Print "======================================="
End Sub
```

**Implementation Tasks:**
```
Day 2:
□ Create DropboxAPI_POC.bas module
□ Copy code above
□ Replace DROPBOX_APP_KEY with your App Key
□ Replace DROPBOX_APP_SECRET with your App Secret
□ Test AuthenticateUser() function
□ Verify token is received

Day 3:
□ Test UploadFile() function
□ Test DownloadFile() function
□ Test CreateFolder() function
□ Test ListFolder() function
□ Document any errors encountered

Deliverable: Working API module with all core functions
```

---

### Phase 3: Test Form (Day 4)

#### Step 3.1: Create Test Form

**Form Name**: `frmDropboxPOC`

**Design:**
```
┌─────────────────────────────────────────────┐
│  Dropbox API - Proof of Concept              │
├─────────────────────────────────────────────┤
│                                               │
│  Authentication:                              │
│  [  Authenticate with Dropbox  ]             │
│  Status: ___________________________         │
│                                               │
│  ─────────────────────────────────────────   │
│                                               │
│  File Operations:                             │
│  [  Upload File  ]  [  Download File  ]      │
│  [  Create Folder  ]  [  List Folder  ]      │
│                                               │
│  ─────────────────────────────────────────   │
│                                               │
│  Test Results:                                │
│  ┌─────────────────────────────────────┐    │
│  │                                      │    │
│  │  [Log output appears here]          │    │
│  │                                      │    │
│  └─────────────────────────────────────┘    │
│                                               │
│  [  Run All Tests  ]  [  Clear Log  ]        │
│                                               │
└─────────────────────────────────────────────┘
```

**Form Code:**
```vba
Option Compare Database
Option Explicit

Private Sub btnAuthenticate_Click()
    Me.txtLog = Me.txtLog & "Authenticating..." & vbCrLf
    
    If DropboxAPI_POC.AuthenticateUser() Then
        Me.txtStatus = "✓ Authenticated"
        Me.txtLog = Me.txtLog & "✓ Authentication successful" & vbCrLf
    Else
        Me.txtStatus = "✗ Not authenticated"
        Me.txtLog = Me.txtLog & "✗ Authentication failed" & vbCrLf
    End If
End Sub

Private Sub btnUpload_Click()
    Dim fd As FileDialog
    Dim localFile As String
    Dim dropboxPath As String
    
    ' File picker
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select file to upload"
    fd.AllowMultiSelect = False
    
    If fd.Show = -1 Then
        localFile = fd.SelectedItems(1)
        dropboxPath = "/TB_CMS_POC/" & Dir(localFile)
        
        Me.txtLog = Me.txtLog & "Uploading: " & localFile & vbCrLf
        
        If DropboxAPI_POC.UploadFile(localFile, dropboxPath) Then
            Me.txtLog = Me.txtLog & "✓ Upload successful" & vbCrLf
        Else
            Me.txtLog = Me.txtLog & "✗ Upload failed" & vbCrLf
        End If
    End If
End Sub

Private Sub btnDownload_Click()
    Dim dropboxPath As String
    Dim localPath As String
    
    dropboxPath = InputBox("Enter Dropbox path (e.g., /TB_CMS_POC/test.pdf):", _
                          "Download File", "/TB_CMS_POC/")
    
    If dropboxPath <> "" Then
        localPath = "C:\Temp\" & Right(dropboxPath, Len(dropboxPath) - InStrRev(dropboxPath, "/"))
        
        Me.txtLog = Me.txtLog & "Downloading: " & dropboxPath & vbCrLf
        
        If DropboxAPI_POC.DownloadFile(dropboxPath, localPath) Then
            Me.txtLog = Me.txtLog & "✓ Download successful: " & localPath & vbCrLf
        Else
            Me.txtLog = Me.txtLog & "✗ Download failed" & vbCrLf
        End If
    End If
End Sub

Private Sub btnCreateFolder_Click()
    Dim folderPath As String
    
    folderPath = InputBox("Enter folder path (e.g., /TB_CMS_POC/TestCase):", _
                         "Create Folder", "/TB_CMS_POC/")
    
    If folderPath <> "" Then
        Me.txtLog = Me.txtLog & "Creating folder: " & folderPath & vbCrLf
        
        If DropboxAPI_POC.CreateFolder(folderPath) Then
            Me.txtLog = Me.txtLog & "✓ Folder created" & vbCrLf
        Else
            Me.txtLog = Me.txtLog & "✗ Folder creation failed" & vbCrLf
        End If
    End If
End Sub

Private Sub btnListFolder_Click()
    Dim folderPath As String
    Dim contents As String
    
    folderPath = InputBox("Enter folder path (e.g., /TB_CMS_POC):", _
                         "List Folder", "/TB_CMS_POC")
    
    If folderPath <> "" Then
        Me.txtLog = Me.txtLog & "Listing folder: " & folderPath & vbCrLf
        
        contents = DropboxAPI_POC.ListFolder(folderPath)
        
        If contents <> "" Then
            Me.txtLog = Me.txtLog & "✓ Folder contents:" & vbCrLf
            Me.txtLog = Me.txtLog & Left(contents, 500) & "..." & vbCrLf
        Else
            Me.txtLog = Me.txtLog & "✗ List folder failed" & vbCrLf
        End If
    End If
End Sub

Private Sub btnRunAllTests_Click()
    Me.txtLog = Me.txtLog & "=== Running All Tests ===" & vbCrLf
    Call DropboxAPI_POC.RunAllTests
    Me.txtLog = Me.txtLog & "=== Tests Complete ===" & vbCrLf
End Sub

Private Sub btnClearLog_Click()
    Me.txtLog = ""
End Sub

Private Sub Form_Load()
    If DropboxAPI_POC.IsAuthenticated() Then
        Me.txtStatus = "✓ Authenticated"
    Else
        Me.txtStatus = "✗ Not authenticated"
    End If
End Sub
```

**Implementation Tasks:**
```
Day 4:
□ Create form with design shown above
□ Add controls:
  - btnAuthenticate (Button)
  - btnUpload (Button)
  - btnDownload (Button)
  - btnCreateFolder (Button)
  - btnListFolder (Button)
  - btnRunAllTests (Button)
  - btnClearLog (Button)
  - txtStatus (TextBox)
  - txtLog (TextBox, Multi-line, Scrollbars)
□ Add form code above
□ Test each button manually

Deliverable: Working test form
```

---

### Phase 4: Testing (Day 5)

#### Step 4.1: Manual Test Scenarios

**Test 1: Authentication**
```
Steps:
1. Open frmDropboxPOC
2. Click "Authenticate with Dropbox"
3. Browser opens to Dropbox
4. Click "Allow" on Dropbox
5. Copy authorization code
6. Paste into Access prompt
7. Click OK

Expected: "Authentication successful" message
Actual: _______________
Pass/Fail: ___________
```

**Test 2: Create Folder**
```
Steps:
1. Click "Create Folder"
2. Enter path: /TB_CMS_POC/TestCase/2023-Smith_John/General
3. Click OK

Expected: "Folder created" message
Actual: _______________
Pass/Fail: ___________
```

**Test 3: Upload File**
```
Steps:
1. Click "Upload File"
2. Select a PDF file (< 5MB)
3. Wait for upload

Expected: "Upload successful" in < 5 seconds
Actual: _______________ seconds
Pass/Fail: ___________
```

**Test 4: List Folder**
```
Steps:
1. Click "List Folder"
2. Enter path: /TB_CMS_POC
3. Click OK

Expected: JSON list of files shown
Actual: _______________
Pass/Fail: ___________
```

**Test 5: Download File**
```
Steps:
1. Click "Download File"
2. Enter path to uploaded file
3. Wait for download

Expected: File downloaded and opened
Actual: _______________
Pass/Fail: ___________
```

**Test 6: Error Handling**
```
Steps:
1. Close internet connection
2. Try to upload file
3. Observe error message

Expected: Clear error message
Actual: _______________
Pass/Fail: ___________
```

**Test 7: Performance**
```
Measure:
- 1MB file upload: _____ seconds
- 1MB file download: _____ seconds
- Create folder: _____ seconds
- List folder (10 files): _____ seconds

Target: < 5 seconds for all operations
Pass/Fail: ___________
```

#### Step 4.2: Performance Metrics

**Create Performance Log:**
```vba
' Add to DropboxAPI_POC.bas

Public Sub LogPerformance(operation As String, duration As Double, fileSize As Long)
    Debug.Print "PERF: " & operation & " | " & _
                Format(duration, "0.00") & "s | " & _
                Format(fileSize / 1024, "0") & "KB"
End Sub
```

**Test Data:**
| Operation | File Size | Duration | Status |
|-----------|-----------|----------|--------|
| Upload | 100 KB | | |
| Upload | 1 MB | | |
| Upload | 5 MB | | |
| Download | 100 KB | | |
| Download | 1 MB | | |
| Download | 5 MB | | |
| Create Folder | N/A | | |
| List Folder (10 files) | N/A | | |

---

### Phase 5: Documentation & Review (Days 6-7)

#### Step 5.1: Document Findings

**Create**: `POC_Results.md`

**Template:**
```markdown
# POC Results - Dropbox API Integration

## Summary
- POC Duration: _____ days
- Tests Completed: _____ / _____
- Success Rate: _____%

## What Worked Well
1. 
2. 
3. 

## Challenges Encountered
1. 
2. 
3. 

## Performance Results
- Average upload time: _____ seconds/MB
- Average download time: _____ seconds/MB
- API reliability: _____%

## Technical Findings
1. OAuth 2.0 implementation: [Easy / Moderate / Difficult]
2. File operations: [Reliable / Needs work]
3. Error handling: [Adequate / Needs improvement]
4. VBA limitations encountered: _____

## Recommendations
1. Proceed with full implementation: [Yes / No / With modifications]
2. Areas needing attention: _____
3. Estimated full implementation timeline: _____ weeks

## Code Quality
- Lines of code: _____
- Reusable for production: _____%
- Refactoring needed: [Yes / No]

## Next Steps
1. 
2. 
3. 
```

#### Step 5.2: Stakeholder Demo

**Prepare 15-minute demonstration:**

**Slide 1: POC Overview**
- What we built
- Timeline (1-2 weeks)
- Success criteria

**Slide 2: Live Demo**
- Authenticate
- Upload document
- Download document
- Create folder structure
- Show Dropbox web interface

**Slide 3: Results**
- Performance metrics
- Success rate
- Challenges encountered

**Slide 4: Recommendations**
- Proceed with full implementation?
- Timeline estimate
- Resource requirements
- Next steps

---

## Success Criteria

### Must Pass (Go/No-Go Decision)
- ✅ OAuth 2.0 authentication works
- ✅ File upload works reliably (100% success rate in test)
- ✅ File download works reliably
- ✅ Folder creation works
- ✅ Performance acceptable (< 5 seconds for 1MB)

### Nice to Have
- ✅ Graceful error handling
- ✅ Performance logging
- ✅ Clean, documented code
- ✅ Positive stakeholder feedback

---

## Timeline Summary

| Phase | Duration | Key Deliverable |
|-------|----------|-----------------|
| **Phase 1: Setup** | Day 1 (4 hours) | Dropbox app registered, dev environment ready |
| **Phase 2: Core API** | Days 2-3 | Working DropboxAPI_POC module |
| **Phase 3: Test Form** | Day 4 | Interactive test interface |
| **Phase 4: Testing** | Day 5 | Test results and metrics |
| **Phase 5: Documentation** | Days 6-7 | POC report and demo |

**Total: 7 days (1-2 weeks with interruptions)**

---

## Deliverables Checklist

### Code
- [ ] DropboxAPI_POC.bas module
- [ ] frmDropboxPOC form
- [ ] Test database (DropboxPOC.accdb)

### Documentation
- [ ] POC Results document
- [ ] Performance metrics
- [ ] Test results
- [ ] Code comments
- [ ] README for running POC

### Presentation
- [ ] Demo slides
- [ ] Video recording of demo (optional)
- [ ] Stakeholder feedback

---

## Risk Mitigation

### Potential Issues & Solutions

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| OAuth flow doesn't work | Low | High | Follow Dropbox documentation exactly, test early |
| VBA limitations | Medium | Medium | Research VBA capabilities beforehand |
| Slow performance | Medium | High | Test with realistic file sizes, optimize if needed |
| Network issues | Medium | Low | Test with/without network, implement basic retry |
| Token expiration | Low | Medium | Test token refresh flow |

---

## Next Steps After POC

### If Successful:
1. **Week 1-2**: Refine API module based on POC learnings
2. **Week 3-4**: Add local cache system
3. **Week 5-6**: Integrate with DocumentManagement.bas
4. **Week 7-8**: UI modifications
5. **Week 9-10**: Testing
6. **Week 11-12**: Migration

### If Issues Found:
1. Document specific blockers
2. Research solutions
3. Revise approach if needed
4. Run focused POC on problem areas
5. Reassess timeline

---

## Resources Needed

### Human Resources
- **Developer**: 1 person, ~32 hours over 1-2 weeks
- **Tester**: Optional, can be same person
- **Stakeholder**: 1 hour for demo/review

### Tools & Access
- MS Access (with VBA)
- Dropbox Business account (trial okay)
- Internet connection
- Test documents (various sizes)

### Budget
- $0 (using Dropbox trial)
- Developer time (internal resource)

---

## Support & Questions

### Dropbox Resources
- API Documentation: https://www.dropbox.com/developers/documentation
- OAuth Guide: https://www.dropbox.com/developers/reference/oauth-guide
- API Explorer: https://dropbox.github.io/dropbox-api-v2-explorer/
- Community: https://www.dropboxforum.com/t5/Developers/bd-p/developers

### VBA Resources
- MSXML2 documentation: Microsoft Docs
- ADODB.Stream: Microsoft Docs
- FileDialog: Microsoft Docs

### Internal Support
- Document blockers immediately
- Share POC code for review
- Request help when stuck > 2 hours

---

## Conclusion

This POC will validate the core technical approach before committing to full development. The focus is on proving:
1. ✅ API integration is feasible in VBA
2. ✅ Performance is acceptable
3. ✅ OAuth flow works smoothly
4. ✅ File operations are reliable

**Expected outcome**: Clear go/no-go decision on full implementation with confidence in approach.

---

**POC Owner**: [Your Name]  
**Start Date**: [Date]  
**Target Completion**: [Date + 7-10 days]  
**Status**: Ready to Begin
