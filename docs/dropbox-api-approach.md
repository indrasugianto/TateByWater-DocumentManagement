# Dropbox API-Only Approach - Technical Overview

**Decision**: Direct Dropbox API integration (no desktop sync)  
**Date**: 2026-01-12  
**Rationale**: Maximum control, no client software dependencies, true cloud-first architecture

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                    MS Access Application                         │
│                                                                  │
│  ┌────────────────┐  ┌──────────────┐  ┌──────────────────┐   │
│  │ User Interface │  │ Document Mgmt│  │   Dropbox API    │   │
│  │   (Forms)      │→ │   (Modified) │→ │    (NEW)         │───┼──→ Dropbox Cloud
│  └────────────────┘  └──────────────┘  └──────────────────┘   │       (REST API)
│                             │                    │              │
│                             ↓                    ↓              │
│  ┌────────────────┐  ┌──────────────┐  ┌──────────────────┐   │
│  │   Database     │  │ Local Cache  │  │  Operation Queue │   │
│  │   (SQL)        │  │  (Temp)      │  │   (Offline)      │   │
│  └────────────────┘  └──────────────┘  └──────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

---

## Key Components

### 1. DropboxAPI.bas - Core API Module

**Purpose**: Direct HTTP communication with Dropbox API

**Key Functions:**
```vba
' Authentication
AuthenticateUser() → OAuth 2.0 flow
RefreshAccessToken() → Auto-refresh before expiry
StoreTokenSecurely() → Encrypted storage

' File Operations
UploadFile(local, dropbox) → Upload with retry
DownloadFile(dropbox, local) → Download with cache
UploadLargeFile(local, dropbox) → Chunked upload for >150MB

' Folder Operations
CreateFolder(path) → Recursive folder creation
DeleteFolder(path) → Delete with confirmation
MoveFolder(from, to) → Move entire folder tree
CopyFolder(from, to) → Copy with structure

' Metadata Operations
GetFileMetadata(path) → File info (size, modified, etc.)
ListFolder(path) → List contents
GetFileVersions(path) → Version history
SearchFiles(query) → Full-text search

' Sharing Operations
CreateSharedLink(path) → Generate secure link
SetLinkExpiration(link, days) → Auto-expire links
RevokeSharedLink(link) → Remove access
```

**Error Handling:**
```vba
' Automatic retry with exponential backoff
' Rate limit detection (429) and throttling
' Token expiration handling (401)
' Network connectivity checks
' Comprehensive logging
```

---

### 2. LocalCache.bas - Performance Optimization

**Purpose**: Minimize API calls, improve response time

**Cache Strategy:**
```vba
' Download on first access
' Cache for 24 hours (configurable)
' Auto-cleanup old files
' Max cache size: 500MB (configurable)
' Content hash verification
```

**Cache Location:**
```
C:\TBCMSCache\
├── {CaseID}_{DocumentType}_{timestamp}.tmp
├── {CaseID}_{DocumentType}_{timestamp}.tmp
└── ...
```

**Key Functions:**
```vba
GetCachedFile(dropboxPath) → Returns local path (downloads if needed)
IsCacheValid(dropboxPath) → Check if cache is fresh
UpdateCache(dropboxPath) → Force cache refresh
ClearCache() → Remove all cached files
GetCacheSize() → Monitor cache usage
CleanOldCache() → Remove files older than threshold
```

**Benefits:**
- ✅ Faster document opening (no re-download)
- ✅ Reduced API usage (stay under rate limits)
- ✅ Better user experience (instant access to recent files)
- ✅ Automatic cache management (no manual cleanup)

---

### 3. DocumentStorageAdapter.bas - Abstraction Layer

**Purpose**: Unified interface for all document operations

**Key Pattern:**
```vba
Public Function SaveDocument(localPath, caseID, docType) As Boolean
    ' 1. Ensure folder exists in Dropbox
    Call EnsureFolderExists(GetDropboxPath(caseID, docType))
    
    ' 2. Upload file via API
    If DropboxAPI.UploadFile(localPath, dropboxPath) Then
        ' 3. Save record to database
        Call SaveDocumentRecord(caseID, docType, dropboxPath)
        
        ' 4. Log operation
        Call LogOperation("Upload", dropboxPath, True)
        
        Return True
    Else
        Return False
    End If
End Function

Public Function OpenDocument(caseID, docType) As Boolean
    ' 1. Get Dropbox path from database
    dropboxPath = GetDocumentPath(caseID, docType)
    
    ' 2. Get cached file (downloads if needed)
    localPath = LocalCache.GetCachedFile(dropboxPath, caseID, docType)
    
    ' 3. Open with default application
    Application.FollowHyperlink localPath
    
    ' 4. Log access
    Call LogOperation("Open", dropboxPath, True)
End Function
```

**All Operations Route Through Adapter:**
- SaveDocument()
- OpenDocument()
- DeleteDocument()
- MoveDocument()
- CopyDocument()
- ShareDocument()
- ListDocuments()

---

### 4. Modified DocumentManagement.bas

**Changes Required:**

| Original Function | Change Required | New Implementation |
|------------------|----------------|-------------------|
| `GetDocumentRootFolder()` | Return Dropbox path | `"/Client Files"` |
| `GetDocumentFolderName()` | Build Dropbox path | `/Client Files/{case}/` |
| `FolderExistsCreate()` | Use API create_folder | `DropboxAPI.CreateFolder()` |
| `SaveScannedFileAs()` | Upload via API | `DropboxAPI.UploadFile()` |
| `OpenDocumentFile()` | Download + cache + open | `LocalCache.GetCachedFile()` |
| `OpenDocumentFolder()` | Open in browser | Open Dropbox web URL |
| `MoveDocumentByCaseStatus()` | Use API move | `DropboxAPI.MoveFolder()` |
| `CopyDocumentToClosedFileScan()` | Use API copy | `DropboxAPI.CopyFolder()` |

**Example Modification:**

**Before** (File System):
```vba
Public Function SaveScannedFileAs(CaseID, DocType, SourceFile, Status) As Boolean
    FolderName = GetDocumentFolderName(CaseID, DocType)
    DestFile = FolderName & GetDocumentFileName(CaseID, DocType)
    
    If FolderExistsCreate(FolderName, True) Then
        FileCopy SourceFile, DestFile
        Call SaveCaseDocument(CaseID, DocType, DestFile)
    End If
End Function
```

**After** (Dropbox API):
```vba
Public Function SaveScannedFileAs(CaseID, DocType, SourceFile, Status) As Boolean
    DropboxPath = GetDropboxFolderPath(CaseID, DocType)
    DestPath = DropboxPath & "/" & GetDocumentFileName(CaseID, DocType)
    
    If EnsureFolderExists(DropboxPath) Then
        If DropboxAPI.UploadFile(SourceFile, DestPath) Then
            Call SaveCaseDocument(CaseID, DocType, DestPath)
            Return True
        End If
    End If
    Return False
End Function
```

---

### 5. DropboxEnhancements.bas - New Features

**Version History:**
```vba
Public Function ShowVersionHistory(CaseID, DocType) As Form
    ' Get all versions from Dropbox
    versions = DropboxAPI.GetFileVersions(GetDropboxPath(CaseID, DocType))
    
    ' Display in popup form
    DoCmd.OpenForm "frmVersionHistory"
    Forms("frmVersionHistory").LoadVersions(versions)
End Function

Public Function RestoreVersion(CaseID, DocType, VersionID) As Boolean
    ' Restore previous version
    path = GetDropboxPath(CaseID, DocType)
    If DropboxAPI.RestoreVersion(path, VersionID) Then
        ' Clear cache to force re-download
        LocalCache.RemoveFromCache(path)
        Return True
    End If
End Function
```

**Document Sharing:**
```vba
Public Function ShareWithClient(CaseID, DocType) As String
    ' Generate secure sharing link
    path = GetDropboxPath(CaseID, DocType)
    link = DropboxAPI.CreateSharedLink(path)
    
    ' Set 30-day expiration
    Call DropboxAPI.SetLinkExpiration(link, 30)
    
    ' Save link to database
    Call SaveSharedLink(CaseID, DocType, link)
    
    ' Copy link to clipboard for easy sharing
    Return link
End Function
```

**Advanced Search:**
```vba
Public Function SearchAllDocuments(searchTerm As String) As Collection
    ' Search across all documents in Dropbox
    results = DropboxAPI.SearchFiles(searchTerm)
    
    ' Filter by accessible cases (security)
    filteredResults = FilterByUserPermissions(results)
    
    Return filteredResults
End Function
```

---

## Database Schema Changes

### New Tables

#### tblDropboxConfiguration
```sql
CREATE TABLE tblDropboxConfiguration (
    ConfigID INT PRIMARY KEY IDENTITY,
    AppKey NVARCHAR(100) NOT NULL,
    AppSecret NVARCHAR(100) NOT NULL, -- Encrypted
    AccessToken NVARCHAR(500), -- Encrypted
    RefreshToken NVARCHAR(500), -- Encrypted
    TokenExpiry DATETIME,
    DropboxRootPath NVARCHAR(200) DEFAULT '/Client Files',
    CacheDirectory NVARCHAR(200) DEFAULT 'C:\TBCMSCache\',
    CacheMaxAgeDays INT DEFAULT 1,
    CacheMaxSizeMB INT DEFAULT 500,
    EnableVersioning BIT DEFAULT 1,
    EnableSharing BIT DEFAULT 1,
    LastSync DATETIME,
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
    DropboxPath NVARCHAR(500) UNIQUE,
    DropboxFileID NVARCHAR(100),
    LocalPath NVARCHAR(500),
    FileSize BIGINT,
    ContentHash NVARCHAR(100), -- For integrity verification
    CachedDate DATETIME DEFAULT GETDATE(),
    LastAccessed DATETIME DEFAULT GETDATE(),
    AccessCount INT DEFAULT 1,
    IsCached BIT DEFAULT 1,
    FOREIGN KEY (CaseID) REFERENCES tblCase(CaseID)
)

CREATE INDEX idx_DropboxPath ON tblDropboxFileCache(DropboxPath)
CREATE INDEX idx_LastAccessed ON tblDropboxFileCache(LastAccessed)
```

#### tblDropboxAuditLog
```sql
CREATE TABLE tblDropboxAuditLog (
    AuditID INT PRIMARY KEY IDENTITY,
    CaseID INT,
    DocumentType NVARCHAR(100),
    DropboxPath NVARCHAR(500),
    Action NVARCHAR(50), -- Upload, Download, Delete, Move, Copy, Share, Restore
    UserID INT,
    UserName NVARCHAR(100),
    ActionDate DATETIME DEFAULT GETDATE(),
    Success BIT,
    ErrorMessage NVARCHAR(MAX),
    FileSize BIGINT,
    DurationMS INT, -- Operation duration in milliseconds
    APICallCount INT, -- Number of API calls made
    FOREIGN KEY (CaseID) REFERENCES tblCase(CaseID),
    FOREIGN KEY (UserID) REFERENCES tblUsers(UserID)
)

CREATE INDEX idx_ActionDate ON tblDropboxAuditLog(ActionDate)
CREATE INDEX idx_CaseID ON tblDropboxAuditLog(CaseID)
```

#### tblDropboxOperationQueue
```sql
CREATE TABLE tblDropboxOperationQueue (
    QueueID INT PRIMARY KEY IDENTITY,
    CaseID INT,
    Operation NVARCHAR(50), -- Upload, Delete, Move, etc.
    SourcePath NVARCHAR(500),
    TargetPath NVARCHAR(500),
    Parameters NVARCHAR(MAX), -- JSON parameters
    Status NVARCHAR(20) DEFAULT 'Pending', -- Pending, Processing, Completed, Failed
    Priority INT DEFAULT 5, -- 1=High, 5=Normal, 10=Low
    CreatedDate DATETIME DEFAULT GETDATE(),
    ProcessedDate DATETIME,
    RetryCount INT DEFAULT 0,
    MaxRetries INT DEFAULT 3,
    ErrorMessage NVARCHAR(MAX),
    FOREIGN KEY (CaseID) REFERENCES tblCase(CaseID)
)

CREATE INDEX idx_Status ON tblDropboxOperationQueue(Status, Priority)
```

### Modified Tables

#### tblCaseDocuments (add columns)
```sql
ALTER TABLE tblCaseDocuments
ADD DropboxFileID NVARCHAR(100) NULL,
    DropboxRev NVARCHAR(100) NULL, -- Revision identifier
    DropboxSharedLink NVARCHAR(500) NULL,
    SharedLinkExpiry DATETIME NULL,
    IsInDropbox BIT DEFAULT 0,
    LastSyncDate DATETIME NULL,
    VersionCount INT DEFAULT 1
```

---

## API Usage Patterns

### Authentication Flow

```
1. User opens MS Access application
   ↓
2. System checks if access token is valid
   ↓
3a. Token Valid → Continue
3b. Token Expired → Refresh token
3c. No Token → OAuth 2.0 flow
   ↓
4. OAuth Flow:
   - Open browser to Dropbox authorization URL
   - User authorizes application
   - User copies authorization code
   - System exchanges code for access + refresh tokens
   - Tokens stored encrypted in database
   ↓
5. System ready for API calls
```

### Upload Flow

```
User clicks "Scan Document"
   ↓
1. Select file from scanner folder
   ↓
2. Get target Dropbox path
   ↓
3. Ensure folder structure exists
   ↓
4. Check file size:
   - Small (<150MB) → Regular upload
   - Large (>150MB) → Chunked upload
   ↓
5. Upload file with progress indicator
   ↓
6. On success:
   - Save file ID to database
   - Update cache (if needed)
   - Log operation
   ↓
7. Show success message
```

### Download Flow

```
User clicks "Open Document"
   ↓
1. Get Dropbox path from database
   ↓
2. Check local cache:
   - Cache hit + valid → Use cached file
   - Cache miss/invalid → Download from Dropbox
   ↓
3. If downloading:
   - Show progress indicator
   - Download to temp cache folder
   - Verify content hash
   - Update cache database
   ↓
4. Open file with default application
   ↓
5. Log access
```

---

## Performance Optimizations

### 1. Caching Strategy
- **Cache recently accessed files** (24-hour default)
- **Cache frequently accessed files** longer (access count based)
- **Pre-cache common files** (e.g., templates, frequently viewed docs)
- **Lazy loading** - Download only when needed
- **Background cleanup** - Remove old cache during idle time

### 2. Batch Operations
```vba
' Instead of individual API calls:
For Each file In files
    DropboxAPI.UploadFile(file) ' N API calls
Next

' Batch multiple operations:
DropboxAPI.BatchUploadFiles(files) ' 1 API call
```

### 3. Rate Limit Management
```vba
' Track API calls per minute
' If approaching limit (300/min):
'   - Queue operations
'   - Add delays between calls
'   - Batch where possible
'   - Show user "Rate limit - please wait"
```

### 4. Parallel Operations (where safe)
```vba
' Upload multiple independent files simultaneously
' Use VBA Timer events for async-like behavior
' Max 3 concurrent operations to avoid overwhelming
```

---

## Error Handling Strategy

### Retry Logic
```vba
Public Function UploadWithRetry(localPath, dropboxPath, maxRetries) As Boolean
    Dim retryCount As Integer
    Dim waitSeconds As Integer
    
    retryCount = 0
    waitSeconds = 1
    
    Do While retryCount < maxRetries
        If DropboxAPI.UploadFile(localPath, dropboxPath) Then
            Return True
        End If
        
        retryCount = retryCount + 1
        
        ' Exponential backoff: 1s, 2s, 4s, 8s, ...
        Application.Wait Now + TimeValue("00:00:" & Format(waitSeconds, "00"))
        waitSeconds = waitSeconds * 2
    Loop
    
    Return False
End Function
```

### Network Connectivity Check
```vba
Public Function IsNetworkAvailable() As Boolean
    On Error Resume Next
    
    ' Try simple Dropbox API call
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://api.dropbox.com/", False
    http.setRequestHeader "Authorization", "Bearer " & GetAccessToken()
    http.send
    
    IsNetworkAvailable = (http.Status = 200)
End Function
```

### Graceful Degradation
```vba
' If network unavailable:
' - Queue operations for later
' - Use cached files (even if stale)
' - Show appropriate messages
' - Auto-retry when network returns
```

---

## Security Considerations

### Token Security
- ✅ Store tokens encrypted (AES-256)
- ✅ Never log tokens
- ✅ Automatic token refresh
- ✅ Tokens stored per-machine (not per-user initially)
- ✅ Admin can revoke tokens remotely

### API Permissions
- ✅ Request minimum necessary scopes
- ✅ Use "files.content.write" (not "files.permanent_delete")
- ✅ Limit to team folder only (not entire Dropbox)
- ✅ Audit all operations

### Data Protection
- ✅ HTTPS for all API calls (enforced by Dropbox)
- ✅ Content hash verification
- ✅ Shared links with expiration dates
- ✅ Audit trail of all access

---

## Deployment Checklist

### Pre-Deployment
- [ ] Dropbox Business account created
- [ ] OAuth app registered
- [ ] API credentials stored securely
- [ ] Development complete and tested
- [ ] Database schema updated
- [ ] Migration script tested
- [ ] Rollback plan documented
- [ ] User training materials ready

### Deployment Day
- [ ] Backup current database
- [ ] Backup current files
- [ ] Deploy updated Access application
- [ ] Run database schema updates
- [ ] Execute migration script
- [ ] Verify file uploads
- [ ] Test all workflows
- [ ] Monitor for errors

### Post-Deployment
- [ ] Monitor API usage (first 7 days)
- [ ] Track error rates
- [ ] Collect user feedback
- [ ] Fine-tune cache settings
- [ ] Optimize performance
- [ ] Schedule follow-up review

---

## Monitoring & Maintenance

### Daily Monitoring
- API error rate (should be < 1%)
- Failed uploads/downloads
- Cache hit rate (target > 70%)
- Average operation time

### Weekly Tasks
- Review audit logs
- Check cache size and cleanup
- Monitor API usage vs limits
- Review user support tickets

### Monthly Tasks
- Performance analysis
- Cost review (Dropbox usage)
- Security audit
- Feature usage analysis
- User satisfaction survey

---

## Benefits of API-Only Approach

### vs. Desktop Sync
| Aspect | API-Only | Desktop Sync |
|--------|----------|--------------|
| **Client Software** | ✅ None needed | ❌ Dropbox app required |
| **Local Disk Space** | ✅ Minimal (cache) | ❌ Full mirror |
| **Control** | ✅ Full control | ⚠️ Limited |
| **Troubleshooting** | ✅ Central logging | ❌ Per-machine |
| **Deployment** | ✅ Easy (just Access app) | ❌ Install on each PC |
| **Performance Monitoring** | ✅ Complete visibility | ⚠️ Limited |
| **Sync Conflicts** | ✅ No conflicts | ❌ Possible |
| **Network Usage** | ✅ On-demand only | ❌ Continuous sync |

### Business Benefits
- ✅ **Lower IT overhead** - No desktop app management
- ✅ **Better scalability** - Add users without local setup
- ✅ **Centralized control** - Manage from one place
- ✅ **Detailed analytics** - Track all operations
- ✅ **Flexible deployment** - Works on any Windows machine

---

## Next Steps

1. **Review technical approach** with development team
2. **Finalize API credentials** and security model
3. **Build prototype** with core functions (upload, download, list)
4. **Test with sample data** (10-20 cases)
5. **Refine caching strategy** based on usage patterns
6. **Complete full development** following plan
7. **Thorough testing** before migration
8. **Execute migration** during scheduled window
9. **Monitor closely** for first 2 weeks
10. **Iterate and improve** based on feedback

---

**Document Version**: 1.0  
**Last Updated**: 2026-01-12  
**Status**: Ready for Development
