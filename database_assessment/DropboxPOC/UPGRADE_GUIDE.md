# Dropbox API POC - Upgrade Guide (v1.0 → v2.0)

## Overview

Version 2.0 includes significant security and reliability improvements:

✅ **Secure token storage** (encrypted in database)  
✅ **Token refresh mechanism** (no re-authentication needed)  
✅ **Configuration in database** (no hardcoded credentials)  
✅ **Input validation** and sanitization  
✅ **Retry logic** with exponential backoff  
✅ **Better error handling** and resource cleanup  
✅ **Activity logging** to database  

---

## 🚀 Quick Start (New Installation)

### Step 1: Run Setup Wizard

In VBA Immediate Window (Ctrl+G), run:

```vba
SetupWizard
```

This interactive wizard will:
1. Create required database tables
2. Configure your Dropbox app credentials
3. Authenticate with Dropbox

**That's it!** You're ready to use the API.

---

## 📋 Manual Setup (Alternative)

If you prefer manual setup:

### Step 1: Create Database Tables

```vba
CreateConfigTables
```

This creates three tables:
- `tblDropboxConfig` - App configuration
- `tblDropboxTokens` - Encrypted token storage
- `tblDropboxLog` - Activity and error logs

### Step 2: Configure Dropbox App Credentials

```vba
SetupDropboxConfig "your_app_key", "your_app_secret"
```

Get credentials from: https://www.dropbox.com/developers/apps

### Step 3: Authenticate

```vba
AuthenticateUser
```

Follow the browser prompts to authorize the app.

---

## 🔄 Migration from v1.0

### Changes Required

1. **Replace Module**
   - Delete old `DropboxAPI_POC` module
   - Import new `DropboxAPI_POC_Updated.bas`
   - Or rename module to `DropboxAPI_POC` if you want to keep the same name

2. **Run Setup (One-Time)**
   ```vba
   SetupWizard
   ```

3. **Update Function Calls** (if needed)
   
   Most function signatures are **unchanged**, so your existing code should work!
   
   ✅ Same signatures:
   ```vba
   UploadFile(localPath, dropboxPath)
   DownloadFile(dropboxPath, localPath)
   CreateFolder(dropboxPath)
   ListFolder(dropboxPath)
   ```
   
   ⚠️ New initialization (optional but recommended):
   ```vba
   ' Add this at the start of your application
   InitializeDropboxAPI
   ```

---

## 📊 Database Schema

### tblDropboxConfig

| Field | Type | Description |
|-------|------|-------------|
| ConfigID | AutoNumber | Primary key |
| ConfigKey | Text(50) | Configuration key (AppKey, AppSecret, RedirectUri) |
| ConfigValue | Text(255) | Configuration value (AppSecret is encrypted) |
| Description | Text(255) | Human-readable description |
| ModifiedDate | DateTime | Last modification timestamp |

### tblDropboxTokens

| Field | Type | Description |
|-------|------|-------------|
| TokenID | AutoNumber | Primary key |
| AccessToken | Text(255) | Encrypted access token |
| RefreshToken | Text(255) | Encrypted refresh token |
| TokenType | Text(50) | Token type (Bearer) |
| ExpiresAt | DateTime | Token expiration timestamp |
| CreatedDate | DateTime | Token creation timestamp |
| IsActive | Yes/No | Whether token is active |

### tblDropboxLog

| Field | Type | Description |
|-------|------|-------------|
| LogID | AutoNumber | Primary key |
| LogDate | DateTime | Log entry timestamp |
| LogLevel | Text(20) | Log level (INFO, ERROR, WARNING) |
| FunctionName | Text(100) | Function that logged the entry |
| ErrorNumber | Long | VBA error number (if applicable) |
| ErrorDescription | Text(255) | Error description (if applicable) |
| Details | Memo | Additional details |

---

## 🔐 Security Improvements

### v1.0 (OLD)
```vba
' ❌ SECURITY RISK: Hardcoded credentials
Private Const DROPBOX_APP_KEY As String = "jbozj8nffezcw9w"
Private Const DROPBOX_APP_SECRET As String = "qjp2rzxzgfhj9qf"

' ❌ Tokens lost when database closes
Private m_AccessToken As String
Private m_RefreshToken As String
```

### v2.0 (NEW)
```vba
' ✅ Credentials stored in database (encrypted)
Private m_AppKey As String        ' Loaded from tblDropboxConfig
Private m_AppSecret As String     ' Loaded encrypted

' ✅ Tokens persisted in database (encrypted)
' Automatically loaded on InitializeDropboxAPI()
```

---

## 🔄 Token Refresh (New Feature)

### How It Works

1. **Token Expiry Tracking**
   - Tokens typically expire after 4 hours
   - System tracks expiry time in `m_TokenExpiry`

2. **Automatic Refresh**
   - Before each API call, checks if token expires soon
   - If < 5 minutes remaining, automatically refreshes
   - Uses refresh token (no user interaction needed!)

3. **Transparent to User**
   - Your code doesn't change
   - Authentication persists across database sessions
   - Only re-authenticate if refresh token expires/revoked

### Usage Example

```vba
' v1.0: User must re-authenticate every session
' ❌ AuthenticateUser() ' Required every time database opens

' v2.0: Tokens loaded automatically
' ✅ InitializeDropboxAPI() ' Loads saved tokens
' ✅ UploadFile(...) ' Automatically refreshes if needed
```

---

## 🔁 Retry Logic (New Feature)

### Built-in Retry with Exponential Backoff

All API functions now automatically retry on transient errors:

- **Rate Limiting (429)**: Waits 2s, 4s, 8s between retries
- **Auth Errors (401)**: Attempts token refresh, then retries
- **Other Errors**: No retry (fail fast)
- **Max Retries**: 3 attempts

### Example Output

```
Attempt 1: Status 429
⚠ Rate limited. Waiting 2 seconds before retry...
Attempt 2: Status 429
⚠ Rate limited. Waiting 4 seconds before retry...
Attempt 3: Status 200
✓ Upload successful!
```

---

## ✅ Input Validation (New Feature)

### v1.0 vs v2.0

**v1.0**: Minimal validation
```vba
If Dir(localFilePath) = "" Then
    MsgBox "File not found"
    Exit Function
End If
```

**v2.0**: Comprehensive validation
```vba
' ✅ Checks:
' - File/folder paths not empty
' - Files exist before upload
' - Directories exist before download
' - Dropbox paths start with /
' - No invalid characters (< > : " | ? *)
' - Token is valid (auto-refresh if needed)
```

---

## 📊 Activity Logging (New Feature)

All API operations are logged to `tblDropboxLog`:

### Success Logs
```vba
LogActivity "UploadFile", "SUCCESS", "Uploaded: /path/to/file.pdf"
```

### Error Logs
```vba
LogError "UploadFile", Err.Number, Err.Description, "File: C:\test.pdf"
```

### Query Examples

```sql
-- View all errors
SELECT * FROM tblDropboxLog 
WHERE LogLevel = 'ERROR' 
ORDER BY LogDate DESC;

-- View recent activity
SELECT * FROM tblDropboxLog 
WHERE LogDate > Date()-7 
ORDER BY LogDate DESC;

-- Upload success rate
SELECT 
    FunctionName,
    COUNT(*) as Total,
    SUM(IIF(LogLevel='SUCCESS',1,0)) as Successful,
    SUM(IIF(LogLevel='ERROR',1,0)) as Failed
FROM tblDropboxLog
WHERE FunctionName = 'UploadFile'
GROUP BY FunctionName;
```

---

## 🧪 Testing

### Run All Tests

```vba
RunAllTests
```

Tests include:
- ✅ Authentication
- ✅ Create folder
- ✅ Upload file
- ✅ List folder
- ✅ Download file

### Individual Tests

```vba
TestAuthentication
TestCreateFolder
TestUpload
TestListFolder
TestDownload
```

---

## 🔧 New Utility Functions

### Check Authentication Status

```vba
If IsAuthenticated() Then
    Debug.Print "✓ Authenticated"
    Debug.Print "Token: " & GetAccessToken()
    Debug.Print "Expires: " & GetTokenExpiry()
Else
    Debug.Print "✗ Not authenticated"
End If
```

### Clear Authentication (Logout)

```vba
ClearAuthentication
' Clears tokens from memory and database
```

### Initialize API (Load Config & Tokens)

```vba
InitializeDropboxAPI
' Loads configuration and tokens
' Auto-refreshes if token expires soon
```

---

## 📝 Code Examples

### Basic Upload/Download

```vba
' Initialize (optional - called automatically on first use)
InitializeDropboxAPI

' Upload file
If UploadFile("C:\Documents\invoice.pdf", "/TB_CMS/2024/Invoice.pdf") Then
    Debug.Print "Upload successful"
End If

' Download file
If DownloadFile("/TB_CMS/2024/Invoice.pdf", "C:\Downloads\invoice.pdf") Then
    Debug.Print "Download successful"
End If
```

### Create Folder Structure

```vba
InitializeDropboxAPI

Dim caseNumber As String
caseNumber = "2024-Smith_John"

' Create case folders
CreateFolder "/TB_CMS/" & caseNumber
CreateFolder "/TB_CMS/" & caseNumber & "/General"
CreateFolder "/TB_CMS/" & caseNumber & "/Legal"
CreateFolder "/TB_CMS/" & caseNumber & "/Medical"
```

### Bulk Upload with Error Handling

```vba
Sub BulkUploadDocuments()
    Dim rs As DAO.Recordset
    Dim successCount As Long
    Dim errorCount As Long
    
    InitializeDropboxAPI
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblDocumentsToUpload")
    
    Do While Not rs.EOF
        If UploadFile(rs!LocalPath, rs!DropboxPath) Then
            successCount = successCount + 1
            rs.Edit
            rs!Uploaded = True
            rs!UploadDate = Now
            rs.Update
        Else
            errorCount = errorCount + 1
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    MsgBox "Upload complete!" & vbCrLf & _
           "Successful: " & successCount & vbCrLf & _
           "Failed: " & errorCount, vbInformation
End Sub
```

---

## 🔒 Encryption Notes

### Current Implementation

The module uses **simple XOR encryption** for token storage:

```vba
Private Function EncryptValue(value As String) As String
    ' Simple XOR encryption
    ' Key = 73 (hardcoded)
```

### ⚠️ Production Recommendation

For production use, implement **Windows Data Protection API (DPAPI)**:

```vba
' Use Windows Credential Manager or DPAPI
' - Machine-specific encryption
' - OS-level key management
' - More secure than XOR
```

**Why not implemented here?**
- Requires Windows API declarations
- More complex code
- POC focuses on workflow, not crypto

**Next Steps:**
- Implement DPAPI for production
- Or use Windows Credential Manager
- Or encrypt entire database file

---

## 🐛 Troubleshooting

### Issue: "Configuration table not found"

**Solution:**
```vba
CreateConfigTables
SetupDropboxConfig "your_app_key", "your_app_secret"
```

### Issue: "Token refresh failed"

**Causes:**
- Refresh token expired (rare, but possible)
- Dropbox app credentials changed
- App was revoked by user

**Solution:**
```vba
ClearAuthentication
AuthenticateUser
```

### Issue: "Upload failed after 3 attempts"

**Check:**
1. Internet connection
2. Dropbox service status
3. File size (< 150MB for basic upload)
4. Dropbox storage quota

**Debug:**
```vba
' Check recent errors
SELECT * FROM tblDropboxLog 
WHERE LogLevel = 'ERROR' 
ORDER BY LogDate DESC;
```

### Issue: "Rate limited"

**Dropbox Rate Limits:**
- Per-user: ~1,000 requests/hour
- Per-app: ~10,000 requests/hour

**Solution:**
- Module automatically retries with backoff
- If persistent, wait 1 hour
- Consider batch operations

---

## 📚 API Reference

### Main Functions

| Function | Description | Returns |
|----------|-------------|---------|
| `InitializeDropboxAPI()` | Load config & tokens, auto-refresh if needed | - |
| `AuthenticateUser()` | OAuth 2.0 authentication flow | Boolean |
| `UploadFile(local, dropbox)` | Upload file to Dropbox | Boolean |
| `DownloadFile(dropbox, local)` | Download file from Dropbox | Boolean |
| `CreateFolder(path)` | Create folder on Dropbox | Boolean |
| `ListFolder(path)` | List folder contents (JSON) | String |

### Setup Functions

| Function | Description |
|----------|-------------|
| `SetupWizard()` | Interactive setup wizard |
| `CreateConfigTables()` | Create database tables |
| `SetupDropboxConfig(key, secret)` | Save configuration |

### Utility Functions

| Function | Description | Returns |
|----------|-------------|---------|
| `IsAuthenticated()` | Check if authenticated | Boolean |
| `GetAccessToken()` | Get token info (debug) | String |
| `GetTokenExpiry()` | Get token expiry | String |
| `ClearAuthentication()` | Logout / clear tokens | - |

### Test Functions

| Function | Description |
|----------|-------------|
| `RunAllTests()` | Run complete test suite |
| `TestAuthentication()` | Test authentication flow |
| `TestCreateFolder()` | Test folder creation |
| `TestUpload()` | Test file upload |
| `TestDownload()` | Test file download |
| `TestListFolder()` | Test folder listing |

---

## 🎯 Best Practices

### 1. Initialize on Startup

Add to your main form's `Form_Load`:

```vba
Private Sub Form_Load()
    ' Initialize Dropbox API
    DropboxAPI_POC_Updated.InitializeDropboxAPI
End Sub
```

### 2. Check Authentication Before Use

```vba
Public Sub UploadCaseDocument(localPath As String, dropboxPath As String)
    If Not IsAuthenticated() Then
        If Not AuthenticateUser() Then
            MsgBox "Authentication failed", vbCritical
            Exit Sub
        End If
    End If
    
    UploadFile localPath, dropboxPath
End Sub
```

### 3. Use Transaction Wrappers

```vba
Public Function UploadWithTracking(docID As Long, localPath As String, dropboxPath As String) As Boolean
    Dim db As DAO.Database
    Set db = CurrentDb
    
    db.Execute "UPDATE tblDocuments SET UploadStatus = 'Uploading' WHERE DocID = " & docID
    
    If UploadFile(localPath, dropboxPath) Then
        db.Execute "UPDATE tblDocuments SET UploadStatus = 'Complete', DropboxPath = '" & dropboxPath & "', UploadDate = Now() WHERE DocID = " & docID
        UploadWithTracking = True
    Else
        db.Execute "UPDATE tblDocuments SET UploadStatus = 'Failed' WHERE DocID = " & docID
        UploadWithTracking = False
    End If
    
    Set db = Nothing
End Function
```

### 4. Monitor Logs

Create a form to view logs:

```vba
' Log Viewer Form - Record Source
SELECT LogDate, LogLevel, FunctionName, ErrorDescription, Details
FROM tblDropboxLog
ORDER BY LogDate DESC;
```

---

## 📈 Version History

### v2.0 (2026-01-14)
- ✅ Secure token storage (encrypted in database)
- ✅ Token refresh mechanism
- ✅ Configuration in database (no hardcoded credentials)
- ✅ Input validation and sanitization
- ✅ Retry logic with exponential backoff
- ✅ Better error handling and resource cleanup
- ✅ Activity logging to database
- ✅ Setup wizard
- ✅ Comprehensive documentation

### v1.0 (2026-01-12)
- Basic OAuth 2.0 authentication
- Upload/download files
- Create folders
- List folder contents
- Simple test suite

---

## 🆘 Support

### Common Questions

**Q: Do I need to re-authenticate every time?**  
A: No! v2.0 saves encrypted tokens. They auto-refresh as needed.

**Q: Can I use the old module name?**  
A: Yes! Just rename `DropboxAPI_POC_Updated` to `DropboxAPI_POC` in the VBA editor.

**Q: Will my existing code break?**  
A: Probably not! Function signatures are unchanged. Just run `SetupWizard()` once.

**Q: How secure is the encryption?**  
A: Basic XOR encryption is suitable for POC. For production, implement DPAPI or use Windows Credential Manager.

**Q: Can I use this with multiple Dropbox accounts?**  
A: Not currently. Module stores one set of tokens. Could be extended to support multiple accounts.

---

## 🎉 You're All Set!

Run the setup wizard and start using the improved Dropbox API:

```vba
SetupWizard
```

Then test it:

```vba
RunAllTests
```

Enjoy the improved reliability and security! 🚀
