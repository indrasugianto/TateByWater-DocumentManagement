# Dropbox for Business Migration Plan

**Project**: TateByWater CMS - Dropbox Integration  
**Date**: 2026-01-12  
**Status**: Planning Phase  
**Goal**: Migrate document management from shared drives to Dropbox for Business while maintaining all current MS Access functionality

---

## Executive Summary

### Current State
- **Storage**: Windows file system (S:\ network drive)
- **Management**: VBA code in MS Access using FileSystemObject
- **Structure**: Hierarchical folders per case with document types
- **Access**: Local network access only

### Target State
- **Storage**: Dropbox for Business cloud storage
- **Management**: VBA code using Dropbox API
- **Structure**: Same hierarchical folder structure (maintain compatibility)
- **Access**: Cloud-based with local sync capabilities
- **Benefits**:
  - ✅ Remote access from anywhere
  - ✅ Built-in versioning (30-day+ history)
  - ✅ Better collaboration and sharing
  - ✅ Automatic backup and sync
  - ✅ Mobile access via Dropbox app
  - ✅ Enhanced security and permissions

---

## Project Phases

### Phase 1: Planning & Design (2 weeks)
- [ ] Dropbox Business account setup
- [ ] API access and app registration
- [ ] Architecture design
- [ ] Code modification planning
- [ ] Testing strategy development

### Phase 2: Development (4-6 weeks)
- [ ] Dropbox API integration module
- [ ] Modify existing VBA code
- [ ] Authentication system
- [ ] Error handling and retry logic
- [ ] Progress indicators for uploads/downloads

### Phase 3: Testing (2-3 weeks)
- [ ] Unit testing of API functions
- [ ] Integration testing
- [ ] User acceptance testing
- [ ] Performance testing
- [ ] Rollback testing

### Phase 4: Migration (1-2 weeks)
- [ ] Backup current data
- [ ] Migrate existing documents
- [ ] Validate migration
- [ ] Update database paths

### Phase 5: Deployment (1 week)
- [ ] Production cutover
- [ ] User training
- [ ] Monitoring and support

**Total Estimated Duration**: 10-14 weeks

---

## Technical Architecture

### Dropbox API Integration Approach

#### Selected Approach: Direct Dropbox API Integration
**Method**: Use Dropbox HTTP API directly via VBA's MSXML2.XMLHTTP for all file operations

**Architecture:**
```
MS Access Application
        │
        ├─→ DropboxAPI.bas (Core API wrapper)
        │   ├─→ OAuth 2.0 Authentication
        │   ├─→ File Upload/Download
        │   ├─→ Folder Operations
        │   └─→ Metadata & Versioning
        │
        ├─→ LocalCache.bas (Temporary file cache)
        │   ├─→ Download on-demand
        │   ├─→ Cache frequently accessed files
        │   └─→ Auto-cleanup old cache
        │
        └─→ DocumentManagement.bas (Modified existing)
            ├─→ All operations via API
            └─→ No local sync dependency
```

**Advantages:**
- ✅ **No desktop app required** - Pure API solution
- ✅ **Full control** - Complete control over all operations
- ✅ **Minimal local storage** - Only temp cache, auto-cleaned
- ✅ **True cloud-first** - Direct cloud access from all machines
- ✅ **Centralized management** - Single source of truth in cloud
- ✅ **Better performance monitoring** - Track all API calls
- ✅ **Flexible deployment** - No client-side software to install

**Challenges & Solutions:**
| Challenge | Solution |
|-----------|----------|
| Network dependency | Implement robust retry logic with exponential backoff |
| Slower file access | Local temp cache for recently accessed files |
| Offline scenarios | Queue operations, sync when online |
| Large file handling | Implement chunked upload/download with progress |
| API rate limits | Batch operations, implement request throttling |

**Why This Approach:**
- Direct API gives maximum flexibility and control
- No need to manage Dropbox desktop installations
- Better for multi-user environments
- Cleaner architecture (single integration point)
- Easier to troubleshoot (all operations logged)
- More scalable (no local disk limitations)

---

## Detailed Implementation Plan

### Phase 1: Planning & Design (Weeks 1-2)

**Focus**: Dropbox API setup, architecture finalization, development environment preparation

#### Week 1: Setup & Analysis

##### 1.1 Dropbox Business Account Setup
```
Tasks:
□ Purchase Dropbox Business plan (Advanced or higher recommended)
□ Create team account
□ Add users with appropriate permissions
□ Set up team folder structure
□ Configure security settings (2FA, SSO if needed)

Deliverable: Configured Dropbox Business account
Time: 2-3 days
```

##### 1.2 Dropbox API Registration
```
Tasks:
□ Register app at https://www.dropbox.com/developers/apps
□ Choose "Scoped access" API
□ Select "Full Dropbox" access (or specific folder)
□ Generate OAuth 2.0 credentials
□ Document App Key and App Secret
□ Set up redirect URIs for OAuth flow

Deliverable: Dropbox API credentials
Time: 1 day
```

##### 1.3 Current System Analysis
```
Tasks:
□ Document all file operations used (create, read, update, delete, move, copy)
□ List all VBA functions that interact with file system
□ Identify critical vs. nice-to-have features
□ Calculate current storage usage
□ Map folder structure requirements

Deliverable: Current system inventory
Time: 2 days
```

#### Week 2: Architecture Design

##### 1.4 Architecture Design Document
```
Tasks:
□ Define Dropbox folder structure (mirror current structure)
□ Design API wrapper module architecture
□ Plan authentication flow (OAuth 2.0)
□ Design error handling strategy
□ Plan for offline/sync scenarios
□ Document migration strategy

Deliverable: Architecture design document
Time: 3 days
```

##### 1.5 Database Schema Updates (if needed)
```
Tasks:
□ Analyze if database schema needs changes
□ Plan for storing Dropbox file IDs (optional)
□ Plan for storing version information
□ Design audit tables for tracking

Deliverable: Database schema change plan
Time: 2 days
```

---

### Phase 2: Development (Weeks 3-8)

#### Week 3-4: Core API Integration

##### 2.1 Create Dropbox API Module
```vba
' New Module: DropboxAPI.bas
' Purpose: Core Dropbox API integration

Option Compare Database
Option Explicit

' API Configuration
Private Const DROPBOX_API_URL As String = "https://api.dropboxapi.com/2/"
Private Const DROPBOX_CONTENT_URL As String = "https://content.dropboxapi.com/2/"
Private Const DROPBOX_APP_KEY As String = "[YOUR_APP_KEY]"
Private Const DROPBOX_APP_SECRET As String = "[YOUR_APP_SECRET]"

' OAuth tokens (stored encrypted in database or registry)
Private m_AccessToken As String

' Core API Functions to Implement:
Public Function AuthenticateUser() As Boolean
Public Function UploadFile(LocalPath As String, DropboxPath As String) As Boolean
Public Function DownloadFile(DropboxPath As String, LocalPath As String) As Boolean
Public Function CreateFolder(DropboxPath As String) As Boolean
Public Function DeleteFile(DropboxPath As String) As Boolean
Public Function MoveFile(FromPath As String, ToPath As String) As Boolean
Public Function CopyFile(FromPath As String, ToPath As String) As Boolean
Public Function ListFolder(DropboxPath As String) As Collection
Public Function GetFileMetadata(DropboxPath As String) As Object
Public Function GetFileVersions(DropboxPath As String) As Collection
Public Function ShareFile(DropboxPath As String) As String
```

**Key Functions:**

```vba
'==============================================================================
' AUTHENTICATION
'==============================================================================
Public Function AuthenticateUser() As Boolean
    ' Implements OAuth 2.0 flow
    ' 1. Show authorization URL to user
    ' 2. User authorizes app
    ' 3. Exchange authorization code for access token
    ' 4. Store access token securely
    
    On Error GoTo ErrHandler
    
    Dim authURL As String
    Dim authCode As String
    Dim http As Object
    Dim response As String
    
    ' Build authorization URL
    authURL = "https://www.dropbox.com/oauth2/authorize?" & _
              "client_id=" & DROPBOX_APP_KEY & _
              "&response_type=code" & _
              "&token_access_type=offline"
    
    ' Open browser for user to authorize
    Application.FollowHyperlink authURL
    
    ' Prompt user to paste authorization code
    authCode = InputBox("Paste the authorization code from Dropbox:", "Dropbox Authorization")
    
    If authCode = "" Then
        AuthenticateUser = False
        Exit Function
    End If
    
    ' Exchange code for access token
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", "https://api.dropbox.com/oauth2/token", False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    http.send "code=" & authCode & _
              "&grant_type=authorization_code" & _
              "&client_id=" & DROPBOX_APP_KEY & _
              "&client_secret=" & DROPBOX_APP_SECRET
    
    If http.Status = 200 Then
        ' Parse response and extract access_token
        response = http.responseText
        m_AccessToken = ExtractAccessToken(response)
        
        ' Store token securely (encrypt and save to database or registry)
        Call StoreAccessToken(m_AccessToken)
        
        AuthenticateUser = True
    Else
        MsgBox "Authentication failed: " & http.responseText
        AuthenticateUser = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Error in AuthenticateUser: " & Err.Description
    AuthenticateUser = False
End Function

'==============================================================================
' UPLOAD FILE
'==============================================================================
Public Function UploadFile(LocalPath As String, DropboxPath As String) As Boolean
    ' Uploads file from local path to Dropbox
    
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim fileBytes() As Byte
    Dim apiArg As String
    
    ' Read file as binary
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' adTypeBinary
    fileStream.Open
    fileStream.LoadFromFile LocalPath
    fileBytes = fileStream.Read
    fileStream.Close
    
    ' Build API argument JSON
    apiArg = "{""path"":""" & DropboxPath & """,""mode"":""overwrite"",""autorename"":false}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", DROPBOX_CONTENT_URL & "files/upload", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Dropbox-API-Arg", apiArg
    http.setRequestHeader "Content-Type", "application/octet-stream"
    
    http.send fileBytes
    
    If http.Status = 200 Then
        UploadFile = True
    Else
        MsgBox "Upload failed: " & http.responseText
        UploadFile = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Error in UploadFile: " & Err.Description
    UploadFile = False
End Function

'==============================================================================
' DOWNLOAD FILE
'==============================================================================
Public Function DownloadFile(DropboxPath As String, LocalPath As String) As Boolean
    ' Downloads file from Dropbox to local path
    
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim apiArg As String
    
    ' Build API argument
    apiArg = "{""path"":""" & DropboxPath & """}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", DROPBOX_CONTENT_URL & "files/download", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Dropbox-API-Arg", apiArg
    
    http.send
    
    If http.Status = 200 Then
        ' Save response to file
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Type = 1 ' adTypeBinary
        fileStream.Open
        fileStream.Write http.responseBody
        fileStream.SaveToFile LocalPath, 2 ' adSaveCreateOverWrite
        fileStream.Close
        
        DownloadFile = True
    Else
        MsgBox "Download failed: " & http.responseText
        DownloadFile = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Error in DownloadFile: " & Err.Description
    DownloadFile = False
End Function

'==============================================================================
' CREATE FOLDER
'==============================================================================
Public Function CreateFolder(DropboxPath As String) As Boolean
    ' Creates folder in Dropbox
    
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    
    ' Build JSON body
    jsonBody = "{""path"":""" & DropboxPath & """,""autorename"":false}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", DROPBOX_API_URL & "files/create_folder_v2", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Content-Type", "application/json"
    
    http.send jsonBody
    
    If http.Status = 200 Then
        CreateFolder = True
    ElseIf InStr(http.responseText, "path/conflict/folder") > 0 Then
        ' Folder already exists - treat as success
        CreateFolder = True
    Else
        MsgBox "Create folder failed: " & http.responseText
        CreateFolder = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Error in CreateFolder: " & Err.Description
    CreateFolder = False
End Function

'==============================================================================
' LIST FOLDER
'==============================================================================
Public Function ListFolder(DropboxPath As String) As Collection
    ' Lists files and folders in specified path
    
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    Dim response As String
    Dim items As Collection
    
    Set items = New Collection
    
    ' Build JSON body
    jsonBody = "{""path"":""" & DropboxPath & """,""recursive"":false}"
    
    ' Make API call
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", DROPBOX_API_URL & "files/list_folder", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Content-Type", "application/json"
    
    http.send jsonBody
    
    If http.Status = 200 Then
        response = http.responseText
        ' Parse JSON response and populate collection
        ' (You'll need a JSON parser or manual parsing)
        Set items = ParseListFolderResponse(response)
    Else
        MsgBox "List folder failed: " & http.responseText
    End If
    
    Set ListFolder = items
    Exit Function
    
ErrHandler:
    MsgBox "Error in ListFolder: " & Err.Description
    Set ListFolder = New Collection
End Function
```

**Tasks:**
```
□ Implement authentication (OAuth 2.0)
□ Implement upload file
□ Implement download file
□ Implement create folder
□ Implement delete file/folder
□ Implement move file/folder
□ Implement copy file/folder
□ Implement list folder contents
□ Implement get file metadata
□ Add error handling and retries
□ Add progress indicators

Deliverable: DropboxAPI.bas module
Time: 2 weeks
```

#### Week 5-6: Create Local Cache & Modify Existing Code

##### 2.2 Create Local Cache Module
```vba
' New Module: LocalCache.bas
' Purpose: Manage temporary local cache of Dropbox files for performance

Option Compare Database
Option Explicit

Private Const CACHE_DIR As String = "C:\TBCMSCache\"
Private Const CACHE_MAX_AGE_HOURS As Long = 24
Private Const CACHE_MAX_SIZE_MB As Long = 500

'==============================================================================
' CACHE MANAGEMENT
'==============================================================================

Public Function GetCachedFile(DropboxPath As String, CaseID As Long, _
                              DocumentType As String) As String
    ' Returns local path to cached file, downloads if needed
    
    On Error GoTo ErrHandler
    
    Dim localPath As String
    Dim cacheInfo As Recordset
    Dim needDownload As Boolean
    
    ' Build local cache path
    localPath = BuildCachePath(CaseID, DocumentType)
    
    ' Check if file is in cache and still valid
    Set cacheInfo = GetCacheInfo(DropboxPath)
    
    If cacheInfo.EOF Then
        ' Not in cache - download it
        needDownload = True
    Else
        ' Check if cache is stale
        If DateDiff("h", cacheInfo("CachedDate"), Now) > CACHE_MAX_AGE_HOURS Then
            needDownload = True
        ElseIf Dir(localPath) = "" Then
            ' Cache record exists but file missing
            needDownload = True
        Else
            ' Verify content hash matches
            If Not VerifyContentHash(localPath, cacheInfo("ContentHash")) Then
                needDownload = True
            Else
                ' Cache is good
                needDownload = False
            End If
        End If
    End If
    
    If needDownload Then
        ' Download from Dropbox
        If DropboxAPI.DownloadFile(DropboxPath, localPath) Then
            ' Update cache info
            Call UpdateCacheInfo(DropboxPath, localPath, CaseID, DocumentType)
        Else
            GetCachedFile = ""
            Exit Function
        End If
    End If
    
    GetCachedFile = localPath
    Exit Function
    
ErrHandler:
    MsgBox "Error in GetCachedFile: " & Err.Description
    GetCachedFile = ""
End Function

Public Sub InitializeCache()
    ' Create cache directory if needed
    If Dir(CACHE_DIR, vbDirectory) = "" Then
        MkDir CACHE_DIR
    End If
    
    ' Clean old cache entries
    Call CleanOldCacheEntries
End Sub

Public Sub CleanOldCacheEntries()
    ' Remove cache entries older than CACHE_MAX_AGE_HOURS
    
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim cutoffDate As Date
    
    cutoffDate = DateAdd("h", -CACHE_MAX_AGE_HOURS, Now)
    
    sql = "SELECT CacheID, LocalPath FROM tblDropboxFileCache " & _
          "WHERE CachedDate < #" & Format(cutoffDate, "yyyy-mm-dd hh:nn:ss") & "#"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    Do Until rs.EOF
        ' Delete physical file
        If Dir(rs("LocalPath")) <> "" Then
            Kill rs("LocalPath")
        End If
        
        ' Delete cache record
        CurrentDb.Execute "DELETE FROM tblDropboxFileCache WHERE CacheID = " & rs("CacheID")
        
        rs.MoveNext
    Loop
    
    rs.Close
    Exit Sub
    
ErrHandler:
    ' Log error but don't stop operation
    Debug.Print "Error in CleanOldCacheEntries: " & Err.Description
End Sub

Public Function GetCacheSize() As Long
    ' Returns total cache size in MB
    
    Dim fso As Object
    Dim folder As Object
    Dim totalSize As Double
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(CACHE_DIR) Then
        Set folder = fso.GetFolder(CACHE_DIR)
        totalSize = folder.Size / 1024 / 1024 ' Convert to MB
    End If
    
    GetCacheSize = CLng(totalSize)
End Function

Public Sub ClearAllCache()
    ' Clears entire cache
    
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(CACHE_DIR) Then
        fso.DeleteFolder CACHE_DIR, True
        MkDir CACHE_DIR
    End If
    
    CurrentDb.Execute "DELETE FROM tblDropboxFileCache"
    
    MsgBox "Cache cleared successfully", vbInformation
End Sub

Private Function BuildCachePath(CaseID As Long, DocumentType As String) As String
    ' Builds local cache file path
    Dim sanitized As String
    sanitized = Replace(DocumentType, "/", "_")
    sanitized = Replace(sanitized, "\", "_")
    sanitized = Replace(sanitized, ":", "_")
    
    BuildCachePath = CACHE_DIR & CaseID & "_" & sanitized & "_" & _
                     Format(Now, "yyyymmddhhnnss") & ".tmp"
End Function

Private Function VerifyContentHash(FilePath As String, ExpectedHash As String) As Boolean
    ' Verify file content matches expected hash
    ' (Implement hash comparison if needed for integrity)
    VerifyContentHash = True ' Simplified for now
End Function

Private Function GetCacheInfo(DropboxPath As String) As DAO.Recordset
    Dim sql As String
    sql = "SELECT * FROM tblDropboxFileCache WHERE DropboxPath = '" & _
          Replace(DropboxPath, "'", "''") & "'"
    Set GetCacheInfo = CurrentDb.OpenRecordset(sql)
End Function

Private Sub UpdateCacheInfo(DropboxPath As String, LocalPath As String, _
                           CaseID As Long, DocumentType As String)
    ' Update or insert cache record
    Dim sql As String
    
    ' Delete old record if exists
    CurrentDb.Execute "DELETE FROM tblDropboxFileCache WHERE DropboxPath = '" & _
                      Replace(DropboxPath, "'", "''") & "'"
    
    ' Insert new record
    sql = "INSERT INTO tblDropboxFileCache (CaseID, DocumentType, DropboxPath, " & _
          "LocalPath, CachedDate, IsCached) VALUES (" & _
          CaseID & ", '" & DocumentType & "', '" & _
          Replace(DropboxPath, "'", "''") & "', '" & _
          Replace(LocalPath, "'", "''") & "', #" & _
          Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, True)"
    
    CurrentDb.Execute sql
End Sub
```

##### 2.3 Create Storage Adapter Module
```vba
' New Module: DocumentStorageAdapter.bas
' Purpose: Abstract storage layer to switch between file system and Dropbox

Option Compare Database
Option Explicit

Public Enum StorageType
    LocalFileSystem = 1
    DropboxCloud = 2
    DropboxSynced = 3
End Enum

' Configuration
Private m_StorageType As StorageType

Public Sub InitializeStorage()
    ' Initialize storage system (always Dropbox API)
    m_StorageType = DropboxCloud
    
    ' Ensure authenticated
    If Not IsDropboxAuthenticated() Then
        If Not DropboxAPI.AuthenticateUser() Then
            MsgBox "Failed to authenticate with Dropbox. Please contact support.", vbCritical
            End
        End If
    End If
    
    ' Initialize local cache
    Call LocalCache.InitializeCache
End Sub

'==============================================================================
' ABSTRACTED FUNCTIONS (Replace existing DocumentManagement functions)
'==============================================================================

Public Function SaveDocument(LocalPath As String, CaseID As Long, _
                            DocumentType As String) As Boolean
    ' Upload document to Dropbox
    On Error GoTo ErrHandler
    
    Dim dropboxPath As String
    Dim folderPath As String
    
    ' Ensure case folder exists
    folderPath = GetDropboxFolderPath(CaseID, DocumentType)
    If Not EnsureFolderExists(folderPath) Then
        SaveDocument = False
        Exit Function
    End If
    
    ' Build full Dropbox path
    dropboxPath = folderPath & "/" & GetDocumentFileName(CaseID, DocumentType, LocalPath)
    
    ' Upload file
    If DropboxAPI.UploadFile(LocalPath, dropboxPath) Then
        ' Save record to database
        If SaveCaseDocumentRecord(CaseID, DocumentType, dropboxPath) Then
            SaveDocument = True
        Else
            SaveDocument = False
        End If
    Else
        SaveDocument = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Error saving document: " & Err.Description, vbExclamation
    SaveDocument = False
End Function

Public Function OpenDocument(CaseID As Long, DocumentType As String) As Boolean
    ' Download and open document from Dropbox
    On Error GoTo ErrHandler
    
    Dim dropboxPath As String
    Dim localPath As String
    
    ' Get Dropbox path from database
    dropboxPath = GetCaseDocumentPath(CaseID, DocumentType)
    
    If dropboxPath = "" Then
        MsgBox "Document not found for this case", vbExclamation
        OpenDocument = False
        Exit Function
    End If
    
    ' Get cached file (downloads if needed)
    localPath = LocalCache.GetCachedFile(dropboxPath, CaseID, DocumentType)
    
    If localPath = "" Then
        MsgBox "Failed to download document", vbExclamation
        OpenDocument = False
        Exit Function
    End If
    
    ' Open with default application
    Application.FollowHyperlink localPath
    
    ' Log access
    Call LogDocumentAccess(CaseID, DocumentType, "Open")
    
    OpenDocument = True
    Exit Function
    
ErrHandler:
    MsgBox "Error opening document: " & Err.Description, vbExclamation
    OpenDocument = False
End Function

Public Function CreateCaseFolder(CaseID As Long, DocumentType As String) As Boolean
    ' Create folder structure in Dropbox
    On Error GoTo ErrHandler
    
    Dim folderPath As String
    
    folderPath = GetDropboxFolderPath(CaseID, DocumentType)
    
    If DropboxAPI.CreateFolder(folderPath) Then
        CreateCaseFolder = True
    Else
        CreateCaseFolder = False
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Error creating folder: " & Err.Description, vbExclamation
    CreateCaseFolder = False
End Function

Private Function EnsureFolderExists(DropboxPath As String) As Boolean
    ' Recursively ensure folder path exists
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    
    parts = Split(DropboxPath, "/")
    currentPath = ""
    
    For i = LBound(parts) To UBound(parts)
        If parts(i) <> "" Then
            currentPath = currentPath & "/" & parts(i)
            ' Try to create (will succeed if exists or newly created)
            Call DropboxAPI.CreateFolder(currentPath)
        End If
    Next i
    
    EnsureFolderExists = True
End Function

Private Function GetDropboxFolderPath(CaseID As Long, DocumentType As String) As String
    ' Build Dropbox folder path
    Dim caseFolder As String
    Dim docFolder As String
    
    ' Get case folder name from database
    caseFolder = GetCaseFolderName(CaseID) ' e.g., "2023-Smith_John"
    
    ' Map document type to folder
    Select Case DocumentType
        Case "General"
            docFolder = "General"
        Case "Client ID"
            docFolder = "ClientID"
        Case "Retainer / Contract"
            docFolder = "Retainer"
        Case "Correspondence: Letters and Emails"
            docFolder = "Correspondence"
        Case "Discovery"
            docFolder = "Discovery"
        Case "Client Invoices"
            docFolder = "Invoices"
        Case "Closed Final"
            docFolder = "ClosedFinal"
        Case Else
            docFolder = "General"
    End Select
    
    GetDropboxFolderPath = "/Client Files/" & caseFolder & "/" & docFolder
End Function
```

**Tasks:**
```
□ Create adapter module for storage abstraction
□ Modify DocumentManagement.bas functions to use adapter
□ Implement file system fallback
□ Update configuration table for storage type
□ Maintain backward compatibility
□ Test switching between storage types

Deliverable: DocumentStorageAdapter.bas + modified DocumentManagement.bas
Time: 1 week
```

##### 2.3 Enhanced Features (Dropbox-Specific)
```vba
' New Module: DropboxEnhancements.bas
' Purpose: Leverage Dropbox-specific features

'==============================================================================
' VERSION HISTORY
'==============================================================================
Public Function GetDocumentVersions(CaseID As Long, DocumentType As String) As Collection
    ' Get version history of a document
    Dim dropboxPath As String
    Dim versions As Collection
    
    dropboxPath = GetDropboxPath(CaseID, DocumentType)
    Set versions = DropboxAPI.GetFileVersions(dropboxPath)
    
    Set GetDocumentVersions = versions
End Function

Public Function RestoreDocumentVersion(CaseID As Long, DocumentType As String, _
                                       VersionID As String) As Boolean
    ' Restore a previous version of document
    ' Implementation here
End Function

'==============================================================================
' SHARING & COLLABORATION
'==============================================================================
Public Function ShareDocumentWithClient(CaseID As Long, DocumentType As String) As String
    ' Generate shareable link for client
    Dim dropboxPath As String
    Dim shareLink As String
    
    dropboxPath = GetDropboxPath(CaseID, DocumentType)
    shareLink = DropboxAPI.ShareFile(dropboxPath)
    
    ShareDocumentWithClient = shareLink
End Function

'==============================================================================
' ADVANCED SEARCH
'==============================================================================
Public Function SearchDocuments(SearchTerm As String) As Collection
    ' Search across all documents in Dropbox
    ' Implementation using Dropbox search API
End Function
```

**Tasks:**
```
□ Implement version history retrieval
□ Implement version restore
□ Implement file sharing links
□ Implement advanced search
□ Add collaboration features
□ Implement file comments (if needed)

Deliverable: DropboxEnhancements.bas
Time: 1 week
```

#### Week 7-8: UI Updates & Configuration

##### 2.4 Update User Interface
```
Tasks:
□ Add "Storage Settings" configuration form
□ Add "Version History" button to document forms
□ Add "Share with Client" button
□ Add progress bars for uploads/downloads
□ Add Dropbox connection status indicator
□ Update help text and tooltips

Deliverable: Updated forms
Time: 3 days
```

##### 2.5 Configuration Management
```
Create new table: tblDropboxConfiguration

Fields:
- ConfigID (PK)
- StorageType (1=FileSystem, 2=Dropbox, 3=Synced)
- DropboxRootPath (e.g., "/Client Files")
- LocalSyncPath (e.g., "C:\Users\...\Dropbox\Client Files")
- AccessToken (encrypted)
- RefreshToken (encrypted)
- LastSync (DateTime)
- EnableVersioning (Boolean)
- EnableAutoSync (Boolean)

Tasks:
□ Create configuration table
□ Create configuration form
□ Implement token storage (encrypted)
□ Implement configuration retrieval functions
□ Add migration flag to track migration status

Deliverable: Configuration system
Time: 2 days
```

##### 2.6 Error Handling & Logging
```vba
' Enhanced error handling for cloud operations

Public Function HandleDropboxError(ErrorCode As Long, ErrorMsg As String) As Boolean
    ' Common error codes:
    ' 401 - Unauthorized (token expired)
    ' 409 - Conflict (file already exists)
    ' 429 - Too many requests (rate limit)
    ' 500 - Server error
    
    Select Case ErrorCode
        Case 401
            ' Token expired - re-authenticate
            If AuthenticateUser() Then
                HandleDropboxError = True ' Retry operation
            Else
                MsgBox "Please re-authenticate with Dropbox", vbCritical
                HandleDropboxError = False
            End If
        
        Case 429
            ' Rate limited - wait and retry
            MsgBox "Too many requests. Waiting 60 seconds...", vbInformation
            Application.Wait Now + TimeValue("00:01:00")
            HandleDropboxError = True ' Retry
        
        Case 409
            ' Conflict - handle based on context
            HandleDropboxError = False
        
        Case Else
            ' Log error and notify user
            Call LogDropboxError(ErrorCode, ErrorMsg)
            MsgBox "Dropbox error: " & ErrorMsg, vbExclamation
            HandleDropboxError = False
    End Select
End Function
```

**Tasks:**
```
□ Implement comprehensive error handling
□ Add retry logic with exponential backoff
□ Create error logging table and functions
□ Add user-friendly error messages
□ Implement offline mode handling

Deliverable: Enhanced error handling
Time: 2 days
```

---

### Phase 3: Testing (Weeks 9-11)

#### Week 9: Unit & Integration Testing

##### 3.1 Unit Tests
```
Test Cases:
□ Authentication flow
□ Upload file (small, medium, large)
□ Download file
□ Create folder (single, nested)
□ Delete file/folder
□ Move file/folder
□ Copy file/folder
□ List folder contents
□ Get file metadata
□ Handle duplicate files
□ Handle special characters in names
□ Handle network interruptions
□ Handle token expiration
□ Handle rate limiting

Deliverable: Test results document
Time: 3 days
```

##### 3.2 Integration Tests
```
Test Workflows:
□ Complete scan workflow (scanner → Dropbox)
□ Open document workflow
□ Case closing workflow (move to _CLOSED)
□ Case reopening workflow
□ Invoice generation and storage
□ Switch between storage types
□ Offline → online transition
□ Bulk document upload

Deliverable: Integration test results
Time: 2 days
```

#### Week 10-11: User Acceptance Testing

##### 3.3 UAT Preparation
```
Tasks:
□ Create test environment with sample data
□ Create test user accounts
□ Prepare test scripts for users
□ Set up test Dropbox team folder
□ Document expected vs. actual results

Deliverable: UAT environment
Time: 2 days
```

##### 3.4 UAT Execution
```
Test Scenarios:
□ Daily document scanning (5-10 documents)
□ Open and review documents
□ Close cases with document movement
□ Search for documents
□ Access version history
□ Share documents with clients
□ Handle errors gracefully
□ Performance (response times)

Participants: 3-5 end users
Duration: 1 week
Deliverable: UAT sign-off
```

##### 3.5 Performance Testing
```
Test Metrics:
□ Upload time (1MB, 10MB, 50MB files)
□ Download time
□ Folder creation time
□ List folder time (10, 100, 1000 files)
□ Concurrent operations
□ Network bandwidth usage

Target Performance:
- Upload: < 2 seconds for 1MB file
- Download: < 2 seconds for 1MB file
- Folder operations: < 1 second
- List folder: < 3 seconds for 100 files

Deliverable: Performance test results
Time: 2 days
```

---

### Phase 4: Migration (Weeks 12-13)

#### Week 12: Data Migration Preparation

##### 4.1 Pre-Migration Checklist
```
□ Complete backup of all current documents
□ Verify backup integrity
□ Document current folder structure
□ Count total files and calculate size
□ Identify any corrupt or inaccessible files
□ Create Dropbox folder structure
□ Set up Dropbox team permissions
□ Prepare migration scripts
□ Schedule migration window (off-hours)
□ Notify all users of migration timeline
```

##### 4.2 Migration Script
```vba
' Module: MigrationUtility.bas
' Purpose: Migrate files from file system to Dropbox

Public Function MigrateAllDocuments() As Boolean
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim totalFiles As Long
    Dim migratedFiles As Long
    Dim failedFiles As Long
    Dim logFile As Integer
    
    ' Open migration log
    logFile = FreeFile
    Open "C:\Migration\migration_log.txt" For Output As #logFile
    
    ' Get all case folders
    Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT CaseID FROM tblCase WHERE Closed = False")
    
    Do Until rs.EOF
        ' Migrate each case
        If MigrateCaseDocuments(rs("CaseID"), logFile) Then
            migratedFiles = migratedFiles + 1
        Else
            failedFiles = failedFiles + 1
        End If
        
        ' Update progress
        DoEvents
        
        rs.MoveNext
    Loop
    
    rs.Close
    Close #logFile
    
    ' Report results
    MsgBox "Migration complete!" & vbCrLf & _
           "Migrated: " & migratedFiles & vbCrLf & _
           "Failed: " & failedFiles, vbInformation
    
    MigrateAllDocuments = (failedFiles = 0)
    Exit Function
    
ErrHandler:
    MsgBox "Migration error: " & Err.Description, vbCritical
    MigrateAllDocuments = False
End Function

Private Function MigrateCaseDocuments(CaseID As Long, LogFile As Integer) As Boolean
    ' Migrate all documents for a specific case
    ' Implementation here
End Function
```

**Tasks:**
```
□ Create migration utility module
□ Test migration script with sample data
□ Create rollback script
□ Prepare migration checklist
□ Schedule migration window

Deliverable: Migration scripts and plan
Time: 3 days
```

#### Week 13: Execute Migration

##### 4.3 Migration Execution
```
Day 1-2: Active Cases
□ Migrate active case documents (highest priority)
□ Verify folder structure
□ Spot-check random files
□ Update database paths
□ Test file access

Day 3-4: Closed Cases
□ Migrate closed case documents
□ Verify migration
□ Update database paths
□ Archive old files (don't delete yet)

Day 5: Verification & Testing
□ Run verification scripts
□ Test all workflows in production
□ Monitor for errors
□ Address any issues
```

##### 4.4 Database Path Updates
```sql
-- Update all document paths to Dropbox format

-- Example: Convert from file system to Dropbox
UPDATE tblCaseDocuments
SET DocumentPath = REPLACE(DocumentPath, 
                          'S:\Client Files\', 
                          '/Client Files/')
WHERE DocumentPath LIKE 'S:\Client Files\%'
```

**Tasks:**
```
□ Execute migration script
□ Verify all files migrated successfully
□ Update database paths
□ Test document access
□ Resolve any migration issues
□ Keep original files as backup (30 days)

Deliverable: Migrated document repository
Time: 1 week
```

---

### Phase 5: Deployment (Week 14)

#### Week 14: Production Cutover

##### 5.1 Production Deployment
```
Day 1: Preparation
□ Final backup of current system
□ Deploy updated Access application
□ Update configuration (set StorageType = Dropbox)
□ Verify Dropbox authentication for all users
□ Enable new features

Day 2: Training
□ Conduct user training sessions
□ Demonstrate new features (versioning, sharing)
□ Provide quick reference guides
□ Answer questions

Day 3-5: Monitoring & Support
□ Monitor system performance
□ Address user issues
□ Fine-tune settings
□ Collect feedback
```

##### 5.2 User Training Materials
```
Create:
□ Quick start guide
□ Video tutorials (5-10 minutes each)
□ FAQ document
□ Troubleshooting guide
□ What's new guide (version history, sharing)
□ Contact information for support

Deliverable: Training materials package
Time: 2 days
```

##### 5.3 Post-Deployment Monitoring
```
Monitor (first 30 days):
□ System errors and exceptions
□ API rate limits and usage
□ User support tickets
□ Performance metrics
□ Storage usage
□ User adoption of new features

Deliverable: Monitoring reports
Time: Ongoing
```

---

## Code Modification Details

### Functions Requiring Changes

From `DocumentManagement.bas` (841 lines):

#### High Priority (Core Functionality)
1. ✅ `GetDocumentRootFolder()` → Get Dropbox root path
2. ✅ `GetDocumentFolderName()` → Build Dropbox path format
3. ✅ `FolderExistsCreate()` → Use Dropbox API create_folder
4. ✅ `SaveScannedFileAs()` → Upload to Dropbox
5. ✅ `OpenDocumentFile()` → Download and open from Dropbox
6. ✅ `OpenDocumentFolder()` → Open Dropbox folder in browser or local sync
7. ✅ `SaveCaseDocument()` → Update with Dropbox path
8. ✅ `GetCaseDocument()` → Retrieve Dropbox path
9. ✅ `MoveDocumentByCaseStatus()` → Use Dropbox move API
10. ✅ `CopyDocumentToClosedFileScan()` → Use Dropbox copy API

#### Medium Priority (Helper Functions)
11. `GetClosedDocumentFolderName()` → Adjust for Dropbox paths
12. `GetIntakeFolderName()` → Adjust for Dropbox paths
13. `GetScannerFolder()` → May remain local or use temp folder
14. `GetClosedFileScanFolderName()` → Adjust for Dropbox paths
15. `GetAllInvoicesFolderName()` → Adjust for Dropbox paths

#### Low Priority (UI Functions)
16. `OpenFileDialog()` → Enhance with Dropbox file picker (optional)
17. `SelectFileDialog()` → Keep for local file selection
18. `GetCaseClosedStatus()` → No changes needed
19. `GetDocumentFileName()` → No changes needed
20. `GetIntakeDocumentFileName()` → No changes needed

---

## Database Schema Changes

### New Tables

#### tblDropboxConfiguration
```sql
CREATE TABLE tblDropboxConfiguration (
    ConfigID INT PRIMARY KEY,
    StorageType INT NOT NULL, -- 1=FileSystem, 2=Dropbox, 3=Synced
    DropboxRootPath NVARCHAR(500),
    LocalSyncPath NVARCHAR(500),
    AppKey NVARCHAR(100),
    AccessToken NVARCHAR(500), -- Encrypted
    RefreshToken NVARCHAR(500), -- Encrypted
    TokenExpiry DATETIME,
    LastSync DATETIME,
    EnableVersioning BIT DEFAULT 1,
    EnableAutoSync BIT DEFAULT 1,
    CreatedDate DATETIME DEFAULT GETDATE(),
    ModifiedDate DATETIME DEFAULT GETDATE()
)
```

#### tblDropboxFileCache
```sql
CREATE TABLE tblDropboxFileCache (
    CacheID INT PRIMARY KEY IDENTITY,
    CaseID INT NOT NULL,
    DocumentType NVARCHAR(100),
    DropboxPath NVARCHAR(500),
    DropboxFileID NVARCHAR(100),
    LocalPath NVARCHAR(500), -- Temp local copy
    LastModified DATETIME,
    FileSize BIGINT,
    ContentHash NVARCHAR(100), -- Dropbox content hash
    IsCached BIT DEFAULT 0,
    CachedDate DATETIME,
    FOREIGN KEY (CaseID) REFERENCES tblCase(CaseID)
)
```

#### tblDropboxAuditLog
```sql
CREATE TABLE tblDropboxAuditLog (
    AuditID INT PRIMARY KEY IDENTITY,
    CaseID INT,
    DocumentType NVARCHAR(100),
    DropboxPath NVARCHAR(500),
    Action NVARCHAR(50), -- Upload, Download, Delete, Move, Copy, Share
    UserID INT,
    ActionDate DATETIME DEFAULT GETDATE(),
    Success BIT,
    ErrorMessage NVARCHAR(MAX),
    FileSize BIGINT,
    Duration INT -- milliseconds
)
```

### Modified Tables

#### tblCaseDocuments (if exists)
```sql
-- Add new columns
ALTER TABLE tblCaseDocuments
ADD DropboxFileID NVARCHAR(100) NULL,
    DropboxRev NVARCHAR(100) NULL,
    DropboxSharedLink NVARCHAR(500) NULL,
    IsInDropbox BIT DEFAULT 0,
    LastSyncDate DATETIME NULL
```

---

## Dropbox Folder Structure

### Proposed Structure (mirrors current)
```
/Client Files/                           [Team folder root]
├── /2023-Smith_John/
│   ├── /General/
│   ├── /ClientID/
│   ├── /Retainer/
│   ├── /Correspondence/
│   ├── /Discovery/
│   ├── /Invoices/
│   └── /ClosedFinal/
├── /2023-Jones_Mary/
│   └── ...
├── /_CLOSED/
│   ├── /2022-Brown_Bob/
│   └── ...
├── /Intakes/                            [Pre-case documents]
│   └── ...
└── /Templates/                          [Optional: Document templates]
    └── ...
```

---

## Security Considerations

### Authentication & Authorization
```
□ Use OAuth 2.0 with refresh tokens (not hardcoded passwords)
□ Encrypt access tokens in database (AES-256)
□ Implement token refresh before expiry
□ Use app-level permissions (not user-level initially)
□ Implement role-based access in application
□ Log all file access attempts
```

### Data Security
```
□ Enable Dropbox encryption at rest (default)
□ Use HTTPS for all API calls (default)
□ Implement audit logging
□ Regular security reviews
□ Comply with legal/ethical requirements for client data
```

### Backup & Disaster Recovery
```
□ Dropbox maintains file versions (30+ days)
□ Keep local backup during transition period (90 days)
□ Document restoration procedures
□ Test disaster recovery scenarios
□ Implement automated backup monitoring
```

---

## Cost Analysis

### Dropbox Business Pricing (2026)
**Dropbox Business Advanced** (recommended):
- **$20/user/month** (billed annually)
- Unlimited storage (or as much as you need)
- 180-day version history
- Advanced sharing controls
- Advanced admin controls
- Full API access

**For 10 users**: $200/month = $2,400/year

**Additional Costs**:
- Development time: 10-14 weeks (see timeline)
- Testing: 2-3 weeks
- Training: 1 week
- Ongoing support: 2-4 hours/month

### Return on Investment
**Benefits**:
- ✅ Remote access (work from anywhere)
- ✅ Automatic versioning (recover deleted files)
- ✅ Better collaboration (real-time sync)
- ✅ Mobile access (iOS/Android apps)
- ✅ Reduced IT infrastructure (no file server maintenance)
- ✅ Automatic backup (no separate backup solution needed)
- ✅ Scalability (add users easily)

**Cost Savings**:
- Eliminate file server maintenance: $200-500/month
- Eliminate separate backup solution: $100-200/month
- Reduce IT support time: 10-20 hours/month

**Break-even**: 6-12 months

---

## Risk Management

### Technical Risks

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|------------|
| API rate limiting | Medium | High | Implement rate limit handling, batch operations |
| Token expiration | High | Medium | Auto-refresh tokens, graceful re-authentication |
| Network connectivity | Medium | High | Offline mode, local caching, retry logic |
| Large file uploads | Medium | Medium | Chunked uploads, progress indicators, resume capability |
| Performance degradation | Low | High | Caching, optimize API calls, monitor performance |
| Data corruption during migration | Low | Critical | Comprehensive backups, validation scripts, rollback plan |

### Business Risks

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|------------|
| User resistance to change | Medium | Medium | Training, gradual rollout, support |
| Data loss during migration | Low | Critical | Backups, validation, pilot testing |
| Downtime during cutover | Medium | High | Off-hours migration, quick rollback plan |
| Budget overrun | Medium | Medium | Phased approach, regular reviews |
| Compliance issues | Low | High | Legal review, security audit |

---

## Success Criteria

### Technical Success Metrics
- ✅ 100% of documents migrated successfully
- ✅ Zero data loss
- ✅ All workflows function correctly
- ✅ API calls complete in < 3 seconds (95th percentile)
- ✅ < 1% error rate on file operations
- ✅ 99.9% uptime (Dropbox SLA)

### User Adoption Metrics
- ✅ 90%+ users successfully using system within 2 weeks
- ✅ < 5 support tickets per user in first month
- ✅ Positive user feedback (> 4/5 rating)
- ✅ Users actively using new features (versioning, sharing)

### Business Metrics
- ✅ ROI positive within 12 months
- ✅ Reduced IT support time by 20%
- ✅ Increased productivity (remote work capability)
- ✅ Improved client satisfaction (document sharing)

---

## Rollback Plan

### If Migration Fails

#### Immediate Rollback (< 24 hours)
```
1. Stop migration script
2. Restore Access application to previous version
3. Set StorageType = FileSystem in configuration
4. Verify file system access works
5. Investigate and document failures
```

#### Post-Migration Rollback (24-72 hours)
```
1. Restore Access application to previous version
2. Update database paths back to file system format
3. Restore any files that were modified
4. Verify all workflows function
5. Communicate to users
```

#### Long-Term Rollback (> 72 hours)
```
1. Schedule maintenance window
2. Restore from comprehensive backup
3. Validate data integrity
4. Test all workflows
5. Conduct post-mortem analysis
6. Plan corrective actions
```

---

## Timeline & Milestones

### Gantt Chart Overview
```
Week 1-2   : [Planning & Design        ]
Week 3-4   : [Core API Development     ]
Week 5-6   : [Code Modifications       ]
Week 7-8   : [UI & Configuration       ]
Week 9     : [Unit Testing             ]
Week 10-11 : [UAT                      ]
Week 12    : [Migration Prep           ]
Week 13    : [Execute Migration        ]
Week 14    : [Deployment & Training    ]
Week 15+   : [Post-Deploy Support      ]
```

### Key Milestones
- ✅ **Week 2**: Architecture design complete
- ✅ **Week 4**: Core API module functional
- ✅ **Week 6**: Code modifications complete
- ✅ **Week 8**: Development complete
- ✅ **Week 11**: UAT sign-off
- ✅ **Week 13**: Migration complete
- ✅ **Week 14**: Production go-live
- ✅ **Week 18**: Post-deployment review

---

## Support & Maintenance

### Post-Deployment Support

#### First 30 Days (Critical Support)
- Daily monitoring of errors and performance
- Immediate response to user issues
- Weekly status reports
- On-call support during business hours

#### 30-90 Days (Active Support)
- Weekly monitoring
- Regular check-ins with users
- Monthly status reports
- Standard support hours

#### 90+ Days (Steady State)
- Monthly monitoring
- Quarterly reviews
- Standard support process

### Maintenance Tasks
```
Monthly:
□ Review API usage and costs
□ Check error logs
□ Review performance metrics
□ Update documentation

Quarterly:
□ Review access tokens and refresh
□ Security audit
□ User feedback survey
□ Feature enhancement planning

Annually:
□ Comprehensive system review
□ Dropbox contract renewal
□ Technology refresh assessment
```

---

## Appendices

### Appendix A: Dropbox API Endpoints

#### Files Endpoints
- `POST /files/upload` - Upload file
- `POST /files/download` - Download file
- `POST /files/delete_v2` - Delete file/folder
- `POST /files/move_v2` - Move file/folder
- `POST /files/copy_v2` - Copy file/folder
- `POST /files/create_folder_v2` - Create folder
- `POST /files/list_folder` - List folder contents
- `POST /files/get_metadata` - Get file/folder metadata
- `POST /files/list_revisions` - Get file version history

#### Sharing Endpoints
- `POST /sharing/create_shared_link_with_settings` - Create share link
- `POST /sharing/list_shared_links` - List share links
- `POST /sharing/revoke_shared_link` - Revoke share link

### Appendix B: VBA JSON Parser
```vba
' Lightweight JSON parser for VBA (if not using external library)
' Include: JsonParser.bas module

Public Function ParseJSON(jsonString As String) As Object
    ' Simple JSON parsing implementation
    ' For production, consider using VBA-JSON library
    ' https://github.com/VBA-tools/VBA-JSON
End Function
```

### Appendix C: Encryption Helper
```vba
' Module: EncryptionHelper.bas
' Purpose: Encrypt/decrypt access tokens

Public Function EncryptToken(token As String) As String
    ' Use Windows DPAPI for encryption
    ' Implementation depends on Windows API calls
End Function

Public Function DecryptToken(encryptedToken As String) As String
    ' Decrypt using Windows DPAPI
End Function
```

---

## Conclusion

This migration plan provides a comprehensive roadmap for moving from file system-based document management to Dropbox for Business. The **hybrid approach** (Phase 1) minimizes risk while providing immediate benefits, with the option to enhance with cloud-native features over time.

### Recommended Next Steps

1. **Review this plan** with stakeholders (IT, management, end users)
2. **Get Dropbox Business account** and test API access
3. **Prioritize features** (MVP vs. nice-to-have)
4. **Assign resources** (developer, tester, project manager)
5. **Create project timeline** with specific dates
6. **Begin Phase 1** (Planning & Design)

### Questions to Resolve

1. What is the total current storage usage?
2. How many active users need Dropbox access?
3. What is the acceptable downtime for migration?
4. Are there any compliance requirements for cloud storage?
5. What is the budget for this project?
6. Who will be the project sponsor/champion?

---

**Document Version**: 1.0  
**Last Updated**: 2026-01-12  
**Next Review**: Before Phase 1 kickoff
