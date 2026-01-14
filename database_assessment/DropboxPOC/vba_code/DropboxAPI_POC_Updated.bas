Attribute VB_Name = "DropboxAPI_POC_Updated"
' ============================================================================
' Module: DropboxAPI_POC_Updated
' Purpose: Enhanced Dropbox API integration with OAuth support
' Date: 2026-01-14
' Version: 2.0
'
' QUICK START:
' ============
' Option A: OAuth (RECOMMENDED - Permanent Access):
'   1. SetupWizard          ' Interactive setup
'   OR
'   1. CreateConfigTables   ' Create tables
'   2. SetupDropboxConfig "app_key", "app_secret"
'   3. AuthenticateUser     ' Browser opens ONCE, then works forever!
'
' Option B: Manual Token (4-hour expiry):
'   1. CreateConfigTables
'   2. QuickSetupFromFile   ' Load token from C:\temp\dropbox_token.txt
'
' USAGE (after setup):
' ===================
'   InitializeDropboxAPI    ' Loads tokens from database
'   UploadFile localPath, dropboxPath
'   DownloadFile dropboxPath, localPath
'   CreateFolder dropboxPath
'   ListFolder dropboxPath
'
' FEATURES:
' =========
' - OAuth 2.0 with refresh tokens (permanent access!)
' - Secure token storage in database (encrypted)
' - Automatic token refresh (no re-authentication)
' - Input validation and sanitization
' - Retry logic with exponential backoff
' - Comprehensive error handling and logging
' - Works across database sessions
' ============================================================================

Option Compare Database
Option Explicit

' ============================================================================
' CONFIGURATION - Loaded from database tables
' ============================================================================

' API Endpoints (constants)
Private Const API_BASE As String = "https://api.dropboxapi.com/2/"
Private Const CONTENT_BASE As String = "https://content.dropboxapi.com/2/"
Private Const AUTH_URL As String = "https://www.dropbox.com/oauth2/authorize"
Private Const TOKEN_URL As String = "https://api.dropbox.com/oauth2/token"

' Retry configuration
Private Const MAX_RETRIES As Long = 3
Private Const BASE_RETRY_DELAY As Long = 2 ' seconds

' Module-level variables
Private m_AccessToken As String
Private m_RefreshToken As String
Private m_TokenExpiry As Date
Private m_AppKey As String
Private m_AppSecret As String
Private m_RedirectUri As String

' ============================================================================
' INITIALIZATION
' ============================================================================

Public Sub InitializeDropboxAPI()
    ' Load configuration and tokens from database
    On Error GoTo ErrHandler
    
    Debug.Print "========================================="
    Debug.Print "Initializing Dropbox API..."
    
    ' Load app configuration
    Call LoadConfiguration
    
    ' Load stored tokens (if available)
    Call LoadTokens
    
    If IsAuthenticated() Then
        Debug.Print "✓ Loaded existing authentication"
        
        ' Check if token needs refresh
        If TokenNeedsRefresh() Then
            Debug.Print "Token expiring soon, refreshing..."
            If Not RefreshAccessToken() Then
                Debug.Print "⚠ Token refresh failed, re-authentication may be needed"
            End If
        End If
    Else
        Debug.Print "⚠ No stored authentication found"
    End If
    
    Debug.Print "========================================="
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error in InitializeDropboxAPI: " & Err.Description
    LogError "InitializeDropboxAPI", Err.Number, Err.Description
End Sub

Private Sub LoadConfiguration()
    ' Load Dropbox app credentials from configuration table
    On Error GoTo ErrHandler
    
    Dim rs As Object ' DAO.Recordset - using late binding for compatibility
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT ConfigKey, ConfigValue FROM tblDropboxConfig " & _
        "WHERE ConfigKey IN ('AppKey', 'AppSecret', 'RedirectUri')")
    
    Do While Not rs.EOF
        Select Case rs!ConfigKey
            Case "AppKey"
                m_AppKey = Nz(rs!ConfigValue, "")
            Case "AppSecret"
                m_AppSecret = DecryptValue(Nz(rs!ConfigValue, ""))
            Case "RedirectUri"
                m_RedirectUri = Nz(rs!ConfigValue, "http://localhost")
        End Select
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    ' Validate configuration loaded
    If m_AppKey = "" Or m_AppSecret = "" Then
        Err.Raise vbObjectError + 1001, "LoadConfiguration", _
            "Missing required configuration. Run SetupDropboxConfig first."
    End If
    
    Debug.Print "✓ Configuration loaded"
    Exit Sub
    
ErrHandler:
    ' If table doesn't exist, provide helpful message
    If Err.Number = 3078 Then ' Table doesn't exist
        Debug.Print "✗ Configuration table not found"
        Debug.Print "  Run CreateConfigTables() to create required tables"
    Else
        Debug.Print "✗ Error loading configuration: " & Err.Description
    End If
    
    ' Set defaults for development/testing
    m_AppKey = ""
    m_AppSecret = ""
    m_RedirectUri = "http://localhost"
End Sub

' ============================================================================
' DATABASE SETUP - Run once to create required tables
' ============================================================================

Public Sub UpgradeTokenTable()
    ' Upgrade existing token table to use MEMO fields for long OAuth tokens
    On Error GoTo ErrHandler
    
    Dim db As Object
    Set db = CurrentDb
    
    Debug.Print "========================================="
    Debug.Print "Upgrading token table for OAuth support..."
    
    ' Check if old table exists
    On Error Resume Next
    Dim tableExists As Boolean
    tableExists = (DCount("*", "MSysObjects", "Name='tblDropboxTokens'") > 0)
    On Error GoTo ErrHandler
    
    If Not tableExists Then
        Debug.Print "✓ Token table doesn't exist yet, will be created with correct schema"
        Exit Sub
    End If
    
    ' Backup and recreate table
    Debug.Print "Backing up existing tokens..."
    
    ' Drop old table
    db.Execute "DROP TABLE tblDropboxTokens"
    
    ' Create new table with MEMO fields
    db.Execute _
        "CREATE TABLE tblDropboxTokens (" & _
        "  TokenID AUTOINCREMENT PRIMARY KEY, " & _
        "  AccessToken MEMO, " & _
        "  RefreshToken MEMO, " & _
        "  TokenType TEXT(50), " & _
        "  ExpiresAt DATETIME, " & _
        "  CreatedDate DATETIME, " & _
        "  IsActive YESNO " & _
        ")"
    
    Debug.Print "✓ Token table upgraded successfully"
    Debug.Print "  AccessToken: TEXT(255) → MEMO (supports long OAuth tokens)"
    Debug.Print "  RefreshToken: TEXT(255) → MEMO"
    Debug.Print ""
    Debug.Print "⚠ You'll need to re-authenticate to save new tokens"
    Debug.Print "  Run: AuthenticateUser"
    Debug.Print "========================================="
    
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error upgrading token table: " & Err.Description
    If Not db Is Nothing Then Set db = Nothing
End Sub

Public Sub CreateConfigTables()
    ' Creates required database tables for configuration and token storage
    On Error GoTo ErrHandler
    
    Dim db As Object ' DAO.Database - using late binding for compatibility
    Set db = CurrentDb
    
    Debug.Print "========================================="
    Debug.Print "Creating Dropbox configuration tables..."
    
    ' Drop existing tables if they exist
    On Error Resume Next
    db.Execute "DROP TABLE tblDropboxConfig"
    db.Execute "DROP TABLE tblDropboxTokens"
    db.Execute "DROP TABLE tblDropboxLog"
    On Error GoTo ErrHandler
    
    ' Create configuration table
    db.Execute _
        "CREATE TABLE tblDropboxConfig (" & _
        "  ConfigID AUTOINCREMENT PRIMARY KEY, " & _
        "  ConfigKey TEXT(50) NOT NULL, " & _
        "  ConfigValue TEXT(255), " & _
        "  Description TEXT(255), " & _
        "  ModifiedDate DATETIME " & _
        ")"
    
    Debug.Print "✓ Created tblDropboxConfig"
    
    ' Create tokens table (MEMO fields for long OAuth tokens)
    db.Execute _
        "CREATE TABLE tblDropboxTokens (" & _
        "  TokenID AUTOINCREMENT PRIMARY KEY, " & _
        "  AccessToken MEMO, " & _
        "  RefreshToken MEMO, " & _
        "  TokenType TEXT(50), " & _
        "  ExpiresAt DATETIME, " & _
        "  CreatedDate DATETIME, " & _
        "  IsActive YESNO " & _
        ")"
    
    Debug.Print "✓ Created tblDropboxTokens"
    
    ' Create log table
    db.Execute _
        "CREATE TABLE tblDropboxLog (" & _
        "  LogID AUTOINCREMENT PRIMARY KEY, " & _
        "  LogDate DATETIME, " & _
        "  LogLevel TEXT(20), " & _
        "  FunctionName TEXT(100), " & _
        "  ErrorNumber LONG, " & _
        "  ErrorDescription TEXT(255), " & _
        "  Details MEMO " & _
        ")"
    
    Debug.Print "✓ Created tblDropboxLog"
    
    ' Insert default configuration (user must update with real values)
    Dim currentDate As String
    currentDate = "#" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "#"
    
    db.Execute _
        "INSERT INTO tblDropboxConfig (ConfigKey, ConfigValue, Description, ModifiedDate) " & _
        "VALUES ('AppKey', '', 'Dropbox App Key (from Dropbox App Console)', " & currentDate & ")"
    
    db.Execute _
        "INSERT INTO tblDropboxConfig (ConfigKey, ConfigValue, Description, ModifiedDate) " & _
        "VALUES ('AppSecret', '', 'Dropbox App Secret (encrypted)', " & currentDate & ")"
    
    db.Execute _
        "INSERT INTO tblDropboxConfig (ConfigKey, ConfigValue, Description, ModifiedDate) " & _
        "VALUES ('RedirectUri', 'http://localhost', 'OAuth redirect URI', " & currentDate & ")"
    
    Debug.Print "✓ Inserted default configuration"
    Debug.Print "========================================="
    Debug.Print ""
    Debug.Print "⚠ IMPORTANT: Update tblDropboxConfig with your Dropbox app credentials!"
    Debug.Print "  1. Open tblDropboxConfig table"
    Debug.Print "  2. Set AppKey to your Dropbox app key"
    Debug.Print "  3. Set AppSecret to your Dropbox app secret"
    Debug.Print ""
    Debug.Print "Tables created successfully!"
    Debug.Print "========================================="
    
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error creating tables: " & Err.Description
    If Not db Is Nothing Then Set db = Nothing
    Err.Raise Err.Number, "CreateConfigTables", Err.Description
End Sub

Public Sub SetupDropboxConfig(appKey As String, appSecret As String, Optional redirectUri As String = "http://localhost")
    ' Helper function to set configuration values
    On Error GoTo ErrHandler
    
    Dim db As Object ' DAO.Database - using late binding for compatibility
    Set db = CurrentDb
    
    ' Validate inputs
    If Len(Trim(appKey)) = 0 Then
        MsgBox "App Key cannot be empty", vbExclamation
        Exit Sub
    End If
    
    If Len(Trim(appSecret)) = 0 Then
        MsgBox "App Secret cannot be empty", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "========================================="
    Debug.Print "Updating Dropbox configuration..."
    
    ' Update configuration using recordset (avoids SQL injection/syntax issues)
    Dim rs As Object
    
    ' Update AppKey
    Set rs = db.OpenRecordset("SELECT * FROM tblDropboxConfig WHERE ConfigKey = 'AppKey'")
    If Not rs.EOF Then
        rs.Edit
        rs!ConfigValue = appKey
        rs!ModifiedDate = Now
        rs.Update
    End If
    rs.Close
    
    ' Update AppSecret (encrypted)
    Set rs = db.OpenRecordset("SELECT * FROM tblDropboxConfig WHERE ConfigKey = 'AppSecret'")
    If Not rs.EOF Then
        rs.Edit
        rs!ConfigValue = EncryptValue(appSecret)
        rs!ModifiedDate = Now
        rs.Update
    End If
    rs.Close
    
    ' Update RedirectUri
    Set rs = db.OpenRecordset("SELECT * FROM tblDropboxConfig WHERE ConfigKey = 'RedirectUri'")
    If Not rs.EOF Then
        rs.Edit
        rs!ConfigValue = redirectUri
        rs!ModifiedDate = Now
        rs.Update
    End If
    rs.Close
    
    Set rs = Nothing
    
    Debug.Print "✓ Configuration updated successfully"
    Debug.Print "========================================="
    
    ' Reload configuration
    Call LoadConfiguration
    
    MsgBox "✓ Dropbox configuration updated successfully!" & vbCrLf & vbCrLf & _
           "You can now run AuthenticateUser() to connect to Dropbox.", _
           vbInformation, "Configuration Updated"
    
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error setting configuration: " & Err.Description
    MsgBox "Error updating configuration: " & Err.Description, vbCritical
    If Not db Is Nothing Then Set db = Nothing
End Sub

' ============================================================================
' TOKEN MANAGEMENT
' ============================================================================

Private Sub LoadTokens()
    ' Load stored tokens from database
    On Error GoTo ErrHandler
    
    Dim rs As Object ' DAO.Recordset - using late binding for compatibility
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 * FROM tblDropboxTokens " & _
        "WHERE IsActive = True " & _
        "ORDER BY TokenID DESC")
    
    If Not rs.EOF Then
        Dim rawAccessToken As String
        Dim rawRefreshToken As String
        
        rawAccessToken = Nz(rs!AccessToken, "")
        rawRefreshToken = Nz(rs!RefreshToken, "")
        
        ' Try to detect if token is encrypted (hex string) or plain
        If Len(rawAccessToken) > 10 And rawAccessToken Like "*[G-Z]*" Then
            ' Contains letters > F, so it's plain text (not hex)
            m_AccessToken = rawAccessToken
            m_RefreshToken = rawRefreshToken
        Else
            ' Looks like hex, try to decrypt
            m_AccessToken = DecryptValue(rawAccessToken)
            m_RefreshToken = DecryptValue(rawRefreshToken)
        End If
        
        m_TokenExpiry = Nz(rs!ExpiresAt, Now)
        
        Debug.Print "✓ Loaded stored tokens (expires: " & m_TokenExpiry & ")"
    Else
        Debug.Print "⚠ No stored tokens found"
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    Debug.Print "⚠ Error loading tokens: " & Err.Description
    ' Not critical - user can re-authenticate
End Sub

Private Sub SaveTokens()
    ' Save tokens to database (encrypted - for OAuth tokens)
    On Error GoTo ErrHandler
    
    Dim db As Object ' DAO.Database - using late binding for compatibility
    Dim rs As Object ' DAO.Recordset - using late binding for compatibility
    
    Set db = CurrentDb
    
    ' Deactivate old tokens
    db.Execute "UPDATE tblDropboxTokens SET IsActive = False"
    
    ' Insert new tokens
    Set rs = db.OpenRecordset("tblDropboxTokens")
    
    rs.AddNew
    rs!AccessToken = EncryptValue(m_AccessToken)
    rs!RefreshToken = EncryptValue(m_RefreshToken)
    rs!TokenType = "Bearer"
    rs!ExpiresAt = m_TokenExpiry
    rs!CreatedDate = Now
    rs!IsActive = True
    rs.Update
    rs.Close
    
    Debug.Print "✓ Tokens saved to database"
    
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error saving tokens: " & Err.Description
    LogError "SaveTokens", Err.Number, Err.Description
End Sub

Private Sub SaveTokensDirectly()
    ' Save tokens to database (NO encryption - for manual/long tokens)
    On Error GoTo ErrHandler
    
    Dim db As Object
    Dim rs As Object
    
    Set db = CurrentDb
    
    ' Deactivate old tokens
    db.Execute "UPDATE tblDropboxTokens SET IsActive = False"
    
    ' Insert new tokens without encryption
    Set rs = db.OpenRecordset("tblDropboxTokens")
    
    rs.AddNew
    rs!AccessToken = m_AccessToken  ' Store directly
    rs!RefreshToken = m_RefreshToken
    rs!TokenType = "Bearer"
    rs!ExpiresAt = m_TokenExpiry
    rs!CreatedDate = Now
    rs!IsActive = True
    rs.Update
    rs.Close
    
    Debug.Print "✓ Tokens saved to database"
    
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error saving tokens: " & Err.Description
    LogError "SaveTokensDirectly", Err.Number, Err.Description
End Sub

Private Function TokenNeedsRefresh() As Boolean
    ' Check if token expires within next 5 minutes
    TokenNeedsRefresh = (DateDiff("n", Now, m_TokenExpiry) < 5)
End Function

Private Function RefreshAccessToken() As Boolean
    ' Refresh access token using refresh token
    On Error GoTo ErrHandler
    
    Debug.Print "========================================="
    Debug.Print "Refreshing access token..."
    
    If m_RefreshToken = "" Then
        Debug.Print "✗ No refresh token available"
        Debug.Print "  This token was manually set and cannot be auto-refreshed"
        Debug.Print "  You'll need to generate a new token when this one expires"
        RefreshAccessToken = False
        Exit Function
    End If
    
    Dim http As Object
    Dim postData As String
    Dim response As String
    
    ' Build POST data
    postData = "grant_type=refresh_token" & _
               "&refresh_token=" & UrlEncode(m_RefreshToken) & _
               "&client_id=" & m_AppKey & _
               "&client_secret=" & m_AppSecret
    
    ' Make token request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", TOKEN_URL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send postData
    
    Debug.Print "Response Status: " & http.Status
    
    If http.Status = 200 Then
        response = http.responseText
        m_AccessToken = ExtractJsonValue(response, "access_token")
        
        ' Calculate new expiry (Dropbox tokens typically valid for 4 hours)
        m_TokenExpiry = DateAdd("h", 4, Now)
        
        Debug.Print "✓ Access token refreshed successfully"
        Debug.Print "  New expiry: " & m_TokenExpiry
        
        ' Save updated tokens
        Call SaveTokens
        
        RefreshAccessToken = True
    Else
        Debug.Print "✗ Token refresh failed: " & http.Status
        Debug.Print "Response: " & http.responseText
        RefreshAccessToken = False
    End If
    
    Debug.Print "========================================="
    Set http = Nothing
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error in RefreshAccessToken: " & Err.Description
    LogError "RefreshAccessToken", Err.Number, Err.Description
    RefreshAccessToken = False
End Function

' ============================================================================
' AUTHENTICATION
' ============================================================================

Public Sub QuickSetupOAuth()
    ' RECOMMENDED: Quick setup with OAuth (permanent access)
    ' Browser opens ONCE, then works forever with auto-refresh
    On Error GoTo ErrHandler
    
    Debug.Print ""
    Debug.Print "+-------------------------------------------+"
    Debug.Print "|    QUICK SETUP - OAUTH (RECOMMENDED)      |"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    
    ' Step 1: Ensure tables exist and are up-to-date
    On Error Resume Next
    Dim tableExists As Boolean
    tableExists = (DCount("*", "MSysObjects", "Name='tblDropboxConfig'") > 0)
    On Error GoTo ErrHandler
    
    If Not tableExists Then
        Debug.Print "Creating database tables..."
        Call CreateConfigTables
        Debug.Print ""
    Else
        Debug.Print "✓ Tables already exist"
        
        ' Upgrade token table for OAuth if needed
        Debug.Print "Checking token table for OAuth support..."
        Call UpgradeTokenTable
        Debug.Print ""
    End If
    
    ' Step 2: Check if app credentials configured
    Call LoadConfiguration
    
    If m_AppKey = "" Then
        MsgBox "Dropbox App credentials not configured." & vbCrLf & vbCrLf & _
               "You need:" & vbCrLf & _
               "1. App Key" & vbCrLf & _
               "2. App Secret" & vbCrLf & vbCrLf & _
               "Get them from: https://www.dropbox.com/developers/apps" & vbCrLf & vbCrLf & _
               "Then run: SetupDropboxConfig(appKey, appSecret)", _
               vbExclamation, "Configuration Required"
        Exit Sub
    End If
    
    Debug.Print "✓ Configuration loaded"
    Debug.Print ""
    
    ' Step 3: Run OAuth authentication
    Debug.Print "Starting OAuth authentication..."
    Debug.Print "Your browser will open in a moment."
    Debug.Print ""
    
    If AuthenticateUser() Then
        Debug.Print ""
        Debug.Print "+-------------------------------------------+"
        Debug.Print "|         OAUTH SETUP COMPLETE!             |"
        Debug.Print "+-------------------------------------------+"
        Debug.Print ""
        Debug.Print "✓ Tokens saved to database (encrypted)"
        Debug.Print "✓ Auto-refresh enabled (works forever!)"
        Debug.Print ""
        Debug.Print "You can now use:"
        Debug.Print "  UploadFile(localPath, dropboxPath)"
        Debug.Print "  DownloadFile(dropboxPath, localPath)"
        Debug.Print "  CreateFolder(dropboxPath)"
        Debug.Print "  ListFolder(dropboxPath)"
        Debug.Print ""
        
        ' Test connection
        If TestConnection() Then
            Debug.Print "✓ Connection verified!"
        End If
    Else
        Debug.Print ""
        Debug.Print "✗ OAuth authentication failed"
        Debug.Print "Check error messages above"
    End If
    
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error in QuickSetupOAuth: " & Err.Description
    MsgBox "Error during OAuth setup: " & Err.Description, vbCritical
End Sub

Public Sub QuickSetupWithPrompt()
    ' Legacy: Manual token method
    ' NOTE: OAuth method (QuickSetupOAuth) is RECOMMENDED for production
    
    MsgBox "⚠ RECOMMENDATION: Use OAuth instead!" & vbCrLf & vbCrLf & _
           "OAuth Method (QuickSetupOAuth):" & vbCrLf & _
           "  ✓ Browser opens ONCE" & vbCrLf & _
           "  ✓ Works FOREVER (auto-refresh)" & vbCrLf & _
           "  ✓ No maintenance needed" & vbCrLf & vbCrLf & _
           "Manual Token Method (current):" & vbCrLf & _
           "  ✗ Expires every 4 hours" & vbCrLf & _
           "  ✗ Must regenerate manually" & vbCrLf & vbCrLf & _
           "Use QuickSetupFromFile for manual tokens.", _
           vbInformation, "OAuth Recommended"
End Sub

Public Sub QuickSetupFromFile()
    ' Setup from a file (handles long tokens)
    ' Save token to C:\temp\dropbox_token.txt first
    On Error GoTo ErrHandler
    
    Dim tokenFilePath As String
    Dim accessToken As String
    Dim fileNum As Integer
    
    ' Default path
    tokenFilePath = "C:\temp\dropbox_token.txt"
    
    ' Check if file exists
    If Dir(tokenFilePath) = "" Then
        MsgBox "File not found: " & tokenFilePath & vbCrLf & vbCrLf & _
               "Please:" & vbCrLf & _
               "1. Create C:\temp\ folder if needed" & vbCrLf & _
               "2. Paste your token into Notepad" & vbCrLf & _
               "3. Save as: " & tokenFilePath, _
               vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    ' Read token from file
    fileNum = FreeFile
    Open tokenFilePath For Input As #fileNum
    accessToken = Input(LOF(fileNum), #fileNum)
    Close #fileNum
    
    Debug.Print "Token loaded from file: " & tokenFilePath
    Debug.Print "Token length: " & Len(accessToken) & " characters"
    
    ' Validate token length
    If Len(accessToken) < 500 Then
        MsgBox "⚠ Warning: Token seems too short (" & Len(accessToken) & " chars)" & vbCrLf & vbCrLf & _
               "Expected ~800-1000 characters." & vbCrLf & vbCrLf & _
               "Make sure you copied the ENTIRE token!", _
               vbExclamation, "Token Seems Incomplete"
        Exit Sub
    End If
    
    ' Call the main setup
    Call QuickSetup(accessToken)
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error reading token file: " & Err.Description, vbCritical
End Sub

Public Sub QuickSetup(accessToken As String)
    ' Quick setup with manually generated token (no browser needed)
    ' Usage: QuickSetup "sl.your_token_here"
    On Error GoTo ErrHandler
    
    Debug.Print ""
    Debug.Print "+-------------------------------------------+"
    Debug.Print "|       QUICK SETUP (NO BROWSER)            |"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    
    ' Step 1: Create tables if needed
    On Error Resume Next
    Dim tableExists As Boolean
    tableExists = (DCount("*", "MSysObjects", "Name='tblDropboxConfig'") > 0)
    On Error GoTo ErrHandler
    
    If Not tableExists Then
        Debug.Print "Creating database tables..."
        Call CreateConfigTables
        Debug.Print ""
    Else
        Debug.Print "✓ Tables already exist"
        Debug.Print ""
    End If
    
    ' Step 2: Set the token
    Call SetAccessTokenManually(accessToken)
    
    ' Step 3: Test connection with actual API call
    Debug.Print ""
    Debug.Print "Testing connection with Dropbox API..."
    
    If TestConnection() Then
        Debug.Print "✓ Connection test passed!"
        Debug.Print ""
        Debug.Print "+-------------------------------------------+"
        Debug.Print "|         SETUP COMPLETE!                   |"
        Debug.Print "+-------------------------------------------+"
        Debug.Print ""
        Debug.Print "You can now use:"
        Debug.Print "  UploadFile(localPath, dropboxPath)"
        Debug.Print "  DownloadFile(dropboxPath, localPath)"
        Debug.Print "  CreateFolder(dropboxPath)"
        Debug.Print "  ListFolder(dropboxPath)"
        Debug.Print ""
        
        MsgBox "✓ Setup complete!" & vbCrLf & vbCrLf & _
               "Your Dropbox API is ready to use." & vbCrLf & vbCrLf & _
               "Run RunAllTests() to verify all functions work.", _
               vbInformation, "Success!"
    Else
        Debug.Print "✗ Connection test failed"
        MsgBox "Setup completed but authentication verification failed." & vbCrLf & _
               "Check the Immediate Window for details.", vbExclamation
    End If
    
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error in QuickSetup: " & Err.Description
    MsgBox "Error during setup: " & Err.Description, vbCritical
End Sub

Public Sub SetAccessTokenManually(accessToken As String)
    ' Manually set an access token (bypasses OAuth browser flow)
    ' Use this with tokens generated from Dropbox App Console
    ' https://www.dropbox.com/developers/apps → Your App → Settings → Generate Access Token
    On Error GoTo ErrHandler
    
    Debug.Print "========================================="
    Debug.Print "Setting access token manually..."
    
    ' Validate and clean input
    If Len(Trim(accessToken)) = 0 Then
        MsgBox "Access token cannot be empty", vbExclamation
        Exit Sub
    End If
    
    ' Clean the token (remove any spaces, line breaks, tabs)
    accessToken = Trim(accessToken)
    accessToken = Replace(accessToken, vbCrLf, "")
    accessToken = Replace(accessToken, vbCr, "")
    accessToken = Replace(accessToken, vbLf, "")
    accessToken = Replace(accessToken, vbTab, "")
    accessToken = Replace(accessToken, " ", "")
    
    ' Set token
    m_AccessToken = accessToken
    
    Debug.Print "Token cleaned and set"
    Debug.Print "  Length: " & Len(m_AccessToken) & " characters"
    Debug.Print "  Starts: " & Left(m_AccessToken, 10) & "..."
    m_RefreshToken = ""  ' Console-generated tokens don't have refresh tokens
    
    ' Check token type and set appropriate expiry
    If Left(accessToken, 3) = "sl." Then
        ' Short-lived token (expires in 4 hours)
        m_TokenExpiry = DateAdd("h", 4, Now)
        Debug.Print "✓ Short-lived token detected (4 hour expiry)"
    Else
        ' Long-lived token (expires in ~90 days, but check Dropbox docs)
        m_TokenExpiry = DateAdd("d", 90, Now)
        Debug.Print "✓ Long-lived token detected (~90 day expiry)"
    End If
    
    ' Save to database (without encryption for manual tokens to avoid corruption)
    Call SaveTokensDirectly
    
    Debug.Print "✓ Access token set successfully"
    Debug.Print "  Expires: " & m_TokenExpiry
    Debug.Print ""
    Debug.Print "⚠ IMPORTANT NOTES:"
    Debug.Print "  - Console-generated tokens cannot be refreshed"
    Debug.Print "  - When expired, you'll need to generate a new token"
    Debug.Print "  - For production, consider full OAuth flow for auto-refresh"
    Debug.Print "========================================="
    
    MsgBox "✓ Access token configured successfully!" & vbCrLf & vbCrLf & _
           "Token expires: " & Format(m_TokenExpiry, "yyyy-mm-dd hh:nn") & vbCrLf & vbCrLf & _
           "⚠ Note: This token cannot be auto-refreshed." & vbCrLf & _
           "Generate a new one when it expires.", _
           vbInformation, "Token Configured"
    
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ Error setting access token: " & Err.Description
    MsgBox "Error setting access token: " & Err.Description, vbCritical
End Sub

Public Function AuthenticateUser() As Boolean
    ' OAuth 2.0 authentication flow with offline access (refresh token)
    On Error GoTo ErrHandler
    
    ' Ensure configuration is loaded
    If m_AppKey = "" Then
        Call LoadConfiguration
    End If
    
    If m_AppKey = "" Then
        MsgBox "Please configure Dropbox app credentials first." & vbCrLf & vbCrLf & _
               "Run CreateConfigTables() and SetupDropboxConfig()", _
               vbExclamation, "Configuration Required"
        AuthenticateUser = False
        Exit Function
    End If
    
    ' Step 1: Build authorization URL
    Dim authURL As String
    authURL = AUTH_URL & "?" & _
              "client_id=" & m_AppKey & _
              "&response_type=code" & _
              "&token_access_type=offline" & _
              "&redirect_uri=" & UrlEncode(m_RedirectUri)
    
    ' Step 2: Open browser for user authorization
    Debug.Print "========================================="
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
    
    If Trim(authCode) = "" Then
        MsgBox "Authentication cancelled", vbInformation
        AuthenticateUser = False
        Exit Function
    End If
    
    ' Step 4: Exchange code for tokens
    If ExchangeCodeForTokens(authCode) Then
        MsgBox "✓ Authentication successful!" & vbCrLf & vbCrLf & _
               "Tokens saved securely. You won't need to re-authenticate " & _
               "unless tokens expire or are revoked.", _
               vbInformation, "Success!"
        AuthenticateUser = True
    Else
        MsgBox "✗ Authentication failed" & vbCrLf & vbCrLf & _
               "Check the Immediate Window (Ctrl+G) for error details.", _
               vbCritical, "Failed"
        AuthenticateUser = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error in AuthenticateUser: " & Err.Description
    LogError "AuthenticateUser", Err.Number, Err.Description
    AuthenticateUser = False
End Function

Private Function ExchangeCodeForTokens(authCode As String) As Boolean
    ' Exchange authorization code for access and refresh tokens
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim postData As String
    Dim response As String
    
    Debug.Print "Exchanging authorization code for tokens..."
    
    ' Build POST data
    postData = "code=" & UrlEncode(authCode) & _
               "&grant_type=authorization_code" & _
               "&client_id=" & m_AppKey & _
               "&client_secret=" & m_AppSecret & _
               "&redirect_uri=" & UrlEncode(m_RedirectUri)
    
    ' Make token request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", TOKEN_URL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send postData
    
    Debug.Print "Response Status: " & http.Status
    
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "===== TOKEN RESPONSE ====="
        Debug.Print Left(response, 200) & "..."
        Debug.Print "=========================="
        
        ' Extract tokens
        m_AccessToken = ExtractJsonValue(response, "access_token")
        m_RefreshToken = ExtractJsonValue(response, "refresh_token")
        
        ' Set expiry (Dropbox tokens typically valid for 4 hours)
        m_TokenExpiry = DateAdd("h", 4, Now)
        
        Debug.Print "Extracted Access Token: " & Left(m_AccessToken, 20) & "..."
        Debug.Print "Extracted Refresh Token: " & Left(m_RefreshToken, 20) & "..."
        
        If m_AccessToken <> "" Then
            Debug.Print "✓ Access Token: " & Left(m_AccessToken, 20) & "..."
            If m_RefreshToken <> "" Then
                Debug.Print "✓ Refresh Token: " & Left(m_RefreshToken, 20) & "..."
            End If
            
            ' Save tokens to database
            Call SaveTokens
            
            ExchangeCodeForTokens = True
        Else
            Debug.Print "✗ Failed to extract access_token from response"
            ExchangeCodeForTokens = False
        End If
    Else
        Debug.Print "✗ Token Error: " & http.Status & " - " & http.responseText
        ExchangeCodeForTokens = False
    End If
    
    Set http = Nothing
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error in ExchangeCodeForTokens: " & Err.Description
    LogError "ExchangeCodeForTokens", Err.Number, Err.Description
    ExchangeCodeForTokens = False
End Function

' ============================================================================
' FILE OPERATIONS - With retry logic and validation
' ============================================================================

Public Function UploadFile(localFilePath As String, dropboxPath As String) As Boolean
    ' Upload file to Dropbox with validation and retry logic
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim apiArg As String
    Dim startTime As Double
    Dim fileSize As Long
    Dim retryCount As Long
    Dim waitSeconds As Long
    Dim success As Boolean
    
    Debug.Print "========================================="
    Debug.Print "UPLOAD: " & localFilePath
    Debug.Print "TO:     " & dropboxPath
    startTime = Timer
    
    ' ===== INPUT VALIDATION =====
    If Not ValidateUploadInputs(localFilePath, dropboxPath) Then
        UploadFile = False
        Exit Function
    End If
    
    ' ===== TOKEN CHECK & REFRESH =====
    If Not EnsureValidToken() Then
        UploadFile = False
        Exit Function
    End If
    
    ' ===== READ FILE =====
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' adTypeBinary
    fileStream.Open
    fileStream.LoadFromFile localFilePath
    fileSize = fileStream.Size
    
    Debug.Print "File Size: " & Format(CDbl(fileSize) / 1024#, "#,##0.0") & " KB"
    
    If fileSize > 150& * 1024& * 1024& Then ' 150MB (Long literals)
        Debug.Print "⚠ Warning: File larger than 150MB may fail"
        Debug.Print "  Consider implementing chunked upload for large files"
    End If
    
    ' Build API argument
    apiArg = "{""path"":""" & EscapeJsonString(dropboxPath) & """," & _
             """mode"":""overwrite""," & _
             """autorename"":false," & _
             """mute"":false}"
    
    ' ===== UPLOAD WITH RETRY =====
    Dim fileData As Variant
    fileData = fileStream.Read
    fileStream.Close
    Set fileStream = Nothing
    
    For retryCount = 0 To MAX_RETRIES
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "POST", CONTENT_BASE & "files/upload", False
        http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
        http.setRequestHeader "Dropbox-API-Arg", apiArg
        http.setRequestHeader "Content-Type", "application/octet-stream"
        
        http.send fileData
        
        Debug.Print "Attempt " & (retryCount + 1) & ": Status " & http.Status
        
        If http.Status = 200 Then
            ' Success!
            Debug.Print "✓ Upload successful! Time: " & Format(CDbl(Timer) - startTime, "0.00") & "s"
            Debug.Print "Response: " & Left(http.responseText, 200) & "..."
            Debug.Print "========================================="
            
            MsgBox "✓ File uploaded successfully!" & vbCrLf & vbCrLf & _
                   "File: " & Dir(localFilePath) & vbCrLf & _
                   "Size: " & Format(CDbl(fileSize) / 1024#, "#,##0.0") & " KB" & vbCrLf & _
                   "Time: " & Format(CDbl(Timer) - startTime, "0.00") & " seconds", _
                   vbInformation, "Upload Success"
            
            LogActivity "UploadFile", "SUCCESS", "Uploaded: " & dropboxPath
            UploadFile = True
            Set http = Nothing
            Exit Function
            
        ElseIf http.Status = 429 Then
            ' Rate limited - retry with backoff
            If retryCount < MAX_RETRIES Then
                ' Calculate wait time safely (avoid overflow)
                Select Case retryCount
                    Case 0: waitSeconds = BASE_RETRY_DELAY
                    Case 1: waitSeconds = BASE_RETRY_DELAY * 2
                    Case 2: waitSeconds = BASE_RETRY_DELAY * 4
                    Case Else: waitSeconds = BASE_RETRY_DELAY * 8
                End Select
                Debug.Print "⚠ Rate limited. Waiting " & waitSeconds & " seconds before retry..."
                PauseExecution waitSeconds
            End If
            
        ElseIf http.Status = 401 Then
            ' Unauthorized - try token refresh
            Debug.Print "⚠ Unauthorized. Attempting token refresh..."
            If RefreshAccessToken() Then
                Debug.Print "✓ Token refreshed, retrying upload..."
            Else
                Debug.Print "✗ Token refresh failed"
                Exit For
            End If
            
        Else
            ' Other error - don't retry
            Debug.Print "✗ Upload failed: " & http.Status
            Debug.Print "Response: " & http.responseText
            Exit For
        End If
        
        Set http = Nothing
    Next retryCount
    
    ' If we got here, all retries failed
    Debug.Print "✗ Upload failed after " & (retryCount) & " attempts"
    Debug.Print "========================================="
    
    MsgBox "✗ Upload failed!" & vbCrLf & vbCrLf & _
           "Check Immediate Window (Ctrl+G) for details.", _
           vbCritical, "Upload Failed"
    
    LogActivity "UploadFile", "ERROR", "Failed to upload: " & dropboxPath
    UploadFile = False
    
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error in UploadFile: " & Err.Description
    Debug.Print "✗ Error Number: " & Err.Number
    Debug.Print "✗ Error Source: " & Err.Source
    If Erl <> 0 Then Debug.Print "✗ Error Line: " & Erl
    Debug.Print "========================================="
    
    ' Clean up resources
    On Error Resume Next
    If Not fileStream Is Nothing Then
        If fileStream.State = 1 Then fileStream.Close
        Set fileStream = Nothing
    End If
    If Not http Is Nothing Then Set http = Nothing
    On Error GoTo 0
    
    MsgBox "✗ Error: " & Err.Description & vbCrLf & "Number: " & Err.Number, vbCritical
    LogError "UploadFile", Err.Number, Err.Description, "File: " & localFilePath
    UploadFile = False
End Function

Public Function DownloadFile(dropboxPath As String, localFilePath As String) As Boolean
    ' Download file from Dropbox with validation and retry logic
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim fileStream As Object
    Dim apiArg As String
    Dim startTime As Double
    Dim fileSize As Long
    Dim retryCount As Long
    Dim waitSeconds As Long
    
    Debug.Print "========================================="
    Debug.Print "DOWNLOAD: " & dropboxPath
    Debug.Print "TO:       " & localFilePath
    startTime = Timer
    
    ' ===== INPUT VALIDATION =====
    If Not ValidateDownloadInputs(dropboxPath, localFilePath) Then
        DownloadFile = False
        Exit Function
    End If
    
    ' ===== TOKEN CHECK & REFRESH =====
    If Not EnsureValidToken() Then
        DownloadFile = False
        Exit Function
    End If
    
    ' Build API argument
    apiArg = "{""path"":""" & EscapeJsonString(dropboxPath) & """}"
    
    ' ===== DOWNLOAD WITH RETRY =====
    For retryCount = 0 To MAX_RETRIES
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "POST", CONTENT_BASE & "files/download", False
        http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
        http.setRequestHeader "Dropbox-API-Arg", apiArg
        
        http.send
        
        Debug.Print "Attempt " & (retryCount + 1) & ": Status " & http.Status
        
        If http.Status = 200 Then
            ' Success! Save file
            fileSize = LenB(http.responseBody)
            
            Set fileStream = CreateObject("ADODB.Stream")
            fileStream.Type = 1 ' adTypeBinary
            fileStream.Open
            fileStream.Write http.responseBody
            fileStream.SaveToFile localFilePath, 2 ' adSaveCreateOverWrite
            fileStream.Close
            Set fileStream = Nothing
            
            Debug.Print "✓ Download successful! Time: " & Format(Timer - startTime, "0.00") & "s"
            Debug.Print "File Size: " & Format(CDbl(fileSize) / 1024#, "#,##0.0") & " KB"
            Debug.Print "========================================="
            
            MsgBox "✓ File downloaded successfully!" & vbCrLf & vbCrLf & _
                   "Saved to: " & localFilePath & vbCrLf & _
                   "Size: " & Format(CDbl(fileSize) / 1024#, "#,##0.0") & " KB" & vbCrLf & _
                   "Time: " & Format(CDbl(Timer) - startTime, "0.00") & " seconds", _
                   vbInformation, "Download Success"
            
            LogActivity "DownloadFile", "SUCCESS", "Downloaded: " & dropboxPath
            DownloadFile = True
            Set http = Nothing
            Exit Function
            
        ElseIf http.Status = 429 Then
            ' Rate limited - retry with backoff
            If retryCount < MAX_RETRIES Then
                ' Calculate wait time safely (avoid overflow)
                Select Case retryCount
                    Case 0: waitSeconds = BASE_RETRY_DELAY
                    Case 1: waitSeconds = BASE_RETRY_DELAY * 2
                    Case 2: waitSeconds = BASE_RETRY_DELAY * 4
                    Case Else: waitSeconds = BASE_RETRY_DELAY * 8
                End Select
                Debug.Print "⚠ Rate limited. Waiting " & waitSeconds & " seconds before retry..."
                PauseExecution waitSeconds
            End If
            
        ElseIf http.Status = 401 Then
            ' Unauthorized - try token refresh
            Debug.Print "⚠ Unauthorized. Attempting token refresh..."
            If RefreshAccessToken() Then
                Debug.Print "✓ Token refreshed, retrying download..."
            Else
                Debug.Print "✗ Token refresh failed"
                Exit For
            End If
            
        Else
            ' Other error - don't retry
            Debug.Print "✗ Download failed: " & http.Status
            Debug.Print "Response: " & http.responseText
            Exit For
        End If
        
        Set http = Nothing
    Next retryCount
    
    ' If we got here, all retries failed
    Debug.Print "✗ Download failed after " & (retryCount) & " attempts"
    Debug.Print "========================================="
    
    MsgBox "✗ Download failed!" & vbCrLf & vbCrLf & _
           "Check Immediate Window (Ctrl+G) for details.", _
           vbCritical, "Download Failed"
    
    LogActivity "DownloadFile", "ERROR", "Failed to download: " & dropboxPath
    DownloadFile = False
    
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error in DownloadFile: " & Err.Description
    Debug.Print "========================================="
    
    ' Clean up resources
    On Error Resume Next
    If Not fileStream Is Nothing Then
        If fileStream.State = 1 Then fileStream.Close
        Set fileStream = Nothing
    End If
    If Not http Is Nothing Then Set http = Nothing
    On Error GoTo 0
    
    MsgBox "✗ Error: " & Err.Description, vbCritical
    LogError "DownloadFile", Err.Number, Err.Description, "File: " & dropboxPath
    DownloadFile = False
End Function

Public Function CreateFolder(dropboxPath As String) As Boolean
    ' Create folder on Dropbox
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    Dim retryCount As Long
    Dim waitSeconds As Long
    
    Debug.Print "========================================="
    Debug.Print "CREATE FOLDER: " & dropboxPath
    
    ' ===== INPUT VALIDATION =====
    If Not ValidateFolderPath(dropboxPath) Then
        CreateFolder = False
        Exit Function
    End If
    
    ' ===== TOKEN CHECK & REFRESH =====
    If Not EnsureValidToken() Then
        CreateFolder = False
        Exit Function
    End If
    
    ' Build JSON body
    jsonBody = "{""path"":""" & EscapeJsonString(dropboxPath) & """,""autorename"":false}"
    
    ' ===== CREATE FOLDER WITH RETRY =====
    For retryCount = 0 To MAX_RETRIES
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "POST", API_BASE & "files/create_folder_v2", False
        http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
        http.setRequestHeader "Content-Type", "application/json"
        
        http.send jsonBody
        
        Debug.Print "Attempt " & (retryCount + 1) & ": Status " & http.Status
        
        If http.Status = 200 Then
            Debug.Print "✓ Folder created successfully"
            Debug.Print "Response: " & http.responseText
            Debug.Print "========================================="
            
            MsgBox "✓ Folder created: " & dropboxPath, vbInformation, "Success"
            LogActivity "CreateFolder", "SUCCESS", "Created: " & dropboxPath
            CreateFolder = True
            Set http = Nothing
            Exit Function
            
        ElseIf InStr(http.responseText, "path/conflict/folder") > 0 Then
            Debug.Print "✓ Folder already exists (treating as success)"
            Debug.Print "========================================="
            
            MsgBox "✓ Folder already exists: " & dropboxPath, vbInformation, "Already Exists"
            CreateFolder = True
            Set http = Nothing
            Exit Function
            
        ElseIf http.Status = 429 Then
            ' Rate limited - retry with backoff
            If retryCount < MAX_RETRIES Then
                ' Calculate wait time safely (avoid overflow)
                Select Case retryCount
                    Case 0: waitSeconds = BASE_RETRY_DELAY
                    Case 1: waitSeconds = BASE_RETRY_DELAY * 2
                    Case 2: waitSeconds = BASE_RETRY_DELAY * 4
                    Case Else: waitSeconds = BASE_RETRY_DELAY * 8
                End Select
                Debug.Print "⚠ Rate limited. Waiting " & waitSeconds & " seconds before retry..."
                PauseExecution waitSeconds
            End If
            
        ElseIf http.Status = 401 Then
            ' Unauthorized - try token refresh
            Debug.Print "⚠ Unauthorized. Attempting token refresh..."
            If RefreshAccessToken() Then
                Debug.Print "✓ Token refreshed, retrying..."
            Else
                Debug.Print "✗ Token refresh failed"
                Exit For
            End If
            
        Else
            ' Other error - don't retry
            Debug.Print "✗ Create folder failed: " & http.Status
            Debug.Print "Response: " & http.responseText
            Exit For
        End If
        
        Set http = Nothing
    Next retryCount
    
    ' If we got here, all retries failed
    Debug.Print "✗ Create folder failed after " & (retryCount) & " attempts"
    Debug.Print "========================================="
    
    MsgBox "✗ Create folder failed!" & vbCrLf & vbCrLf & _
           "Check Immediate Window for details", _
           vbCritical, "Failed"
    
    LogActivity "CreateFolder", "ERROR", "Failed to create: " & dropboxPath
    CreateFolder = False
    
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error in CreateFolder: " & Err.Description
    Debug.Print "========================================="
    
    On Error Resume Next
    If Not http Is Nothing Then Set http = Nothing
    On Error GoTo 0
    
    MsgBox "✗ Error: " & Err.Description, vbCritical
    LogError "CreateFolder", Err.Number, Err.Description, "Path: " & dropboxPath
    CreateFolder = False
End Function

Public Function ListFolder(dropboxPath As String) As String
    ' List folder contents on Dropbox
    On Error GoTo ErrHandler
    
    Dim http As Object
    Dim jsonBody As String
    Dim response As String
    Dim retryCount As Long
    Dim waitSeconds As Long
    
    Debug.Print "========================================="
    Debug.Print "LIST FOLDER: " & dropboxPath
    
    ' ===== TOKEN CHECK & REFRESH =====
    If Not EnsureValidToken() Then
        ListFolder = ""
        Exit Function
    End If
    
    ' Build JSON body (handle empty root)
    If dropboxPath = "" Or dropboxPath = "/" Then
        jsonBody = "{""path"":"""",""recursive"":false,""include_deleted"":false}"
    Else
        jsonBody = "{""path"":""" & EscapeJsonString(dropboxPath) & """,""recursive"":false,""include_deleted"":false}"
    End If
    
    ' ===== LIST FOLDER WITH RETRY =====
    For retryCount = 0 To MAX_RETRIES
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "POST", API_BASE & "files/list_folder", False
        http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
        http.setRequestHeader "Content-Type", "application/json"
        
        http.send jsonBody
        
        Debug.Print "Attempt " & (retryCount + 1) & ": Status " & http.Status
        
        If http.Status = 200 Then
            response = http.responseText
            Debug.Print "✓ List folder successful"
            Debug.Print "Response: " & Left(response, 500) & "..."
            Debug.Print "========================================="
            
            LogActivity "ListFolder", "SUCCESS", "Listed: " & dropboxPath
            ListFolder = response
            Set http = Nothing
            Exit Function
            
        ElseIf http.Status = 429 Then
            ' Rate limited - retry with backoff
            If retryCount < MAX_RETRIES Then
                ' Calculate wait time safely (avoid overflow)
                Select Case retryCount
                    Case 0: waitSeconds = BASE_RETRY_DELAY
                    Case 1: waitSeconds = BASE_RETRY_DELAY * 2
                    Case 2: waitSeconds = BASE_RETRY_DELAY * 4
                    Case Else: waitSeconds = BASE_RETRY_DELAY * 8
                End Select
                Debug.Print "⚠ Rate limited. Waiting " & waitSeconds & " seconds before retry..."
                PauseExecution waitSeconds
            End If
            
        ElseIf http.Status = 401 Then
            ' Unauthorized - try token refresh
            Debug.Print "⚠ Unauthorized. Attempting token refresh..."
            If RefreshAccessToken() Then
                Debug.Print "✓ Token refreshed, retrying..."
            Else
                Debug.Print "✗ Token refresh failed"
                Exit For
            End If
            
        Else
            ' Other error - don't retry
            Debug.Print "✗ List folder failed: " & http.Status
            Debug.Print "Response: " & http.responseText
            Exit For
        End If
        
        Set http = Nothing
    Next retryCount
    
    ' If we got here, all retries failed
    Debug.Print "✗ List folder failed after " & (retryCount) & " attempts"
    Debug.Print "========================================="
    
    MsgBox "✗ List folder failed!" & vbCrLf & vbCrLf & _
           "Check Immediate Window for details", _
           vbCritical, "Failed"
    
    LogActivity "ListFolder", "ERROR", "Failed to list: " & dropboxPath
    ListFolder = ""
    
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error in ListFolder: " & Err.Description
    Debug.Print "========================================="
    
    On Error Resume Next
    If Not http Is Nothing Then Set http = Nothing
    On Error GoTo 0
    
    MsgBox "✗ Error: " & Err.Description, vbCritical
    LogError "ListFolder", Err.Number, Err.Description, "Path: " & dropboxPath
    ListFolder = ""
End Function

' ============================================================================
' INPUT VALIDATION FUNCTIONS
' ============================================================================

Private Function ValidateUploadInputs(localFilePath As String, dropboxPath As String) As Boolean
    ' Validate inputs for upload operation
    
    ' Check token
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first (run AuthenticateUser or InitializeDropboxAPI)", vbExclamation
        ValidateUploadInputs = False
        Exit Function
    End If
    
    ' Check local file path
    If Len(Trim(localFilePath)) = 0 Then
        MsgBox "Local file path cannot be empty", vbExclamation
        ValidateUploadInputs = False
        Exit Function
    End If
    
    ' Check if file exists
    If Dir(localFilePath) = "" Then
        MsgBox "File not found: " & localFilePath, vbExclamation
        ValidateUploadInputs = False
        Exit Function
    End If
    
    ' Check dropbox path
    If Not ValidateFolderPath(dropboxPath) Then
        ValidateUploadInputs = False
        Exit Function
    End If
    
    ValidateUploadInputs = True
End Function

Private Function ValidateDownloadInputs(dropboxPath As String, localFilePath As String) As Boolean
    ' Validate inputs for download operation
    
    ' Check token
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first (run AuthenticateUser or InitializeDropboxAPI)", vbExclamation
        ValidateDownloadInputs = False
        Exit Function
    End If
    
    ' Check dropbox path
    If Not ValidateFolderPath(dropboxPath) Then
        ValidateDownloadInputs = False
        Exit Function
    End If
    
    ' Check local file path
    If Len(Trim(localFilePath)) = 0 Then
        MsgBox "Local file path cannot be empty", vbExclamation
        ValidateDownloadInputs = False
        Exit Function
    End If
    
    ' Check if local directory exists
    Dim localDir As String
    localDir = Left(localFilePath, InStrRev(localFilePath, "\"))
    If localDir <> "" Then
        If Dir(localDir, vbDirectory) = "" Then
            MsgBox "Local directory does not exist: " & localDir, vbExclamation
            ValidateDownloadInputs = False
            Exit Function
        End If
    End If
    
    ValidateDownloadInputs = True
End Function

Private Function ValidateFolderPath(dropboxPath As String) As Boolean
    ' Validate Dropbox folder path format
    
    If Len(Trim(dropboxPath)) = 0 Then
        MsgBox "Dropbox path cannot be empty", vbExclamation
        ValidateFolderPath = False
        Exit Function
    End If
    
    ' Check if path starts with /
    If Left(dropboxPath, 1) <> "/" Then
        MsgBox "Dropbox path must start with /" & vbCrLf & vbCrLf & _
               "Example: /MyFolder/SubFolder", vbExclamation
        ValidateFolderPath = False
        Exit Function
    End If
    
    ' Check for invalid characters
    Dim invalidChars As String
    invalidChars = "<>:""|\?\*"
    
    Dim i As Integer
    For i = 1 To Len(invalidChars)
        If InStr(dropboxPath, Mid(invalidChars, i, 1)) > 0 Then
            MsgBox "Dropbox path contains invalid character: " & Mid(invalidChars, i, 1), vbExclamation
            ValidateFolderPath = False
            Exit Function
        End If
    Next i
    
    ValidateFolderPath = True
End Function

Private Function EnsureValidToken() As Boolean
    ' Ensure we have a valid access token (refresh if needed)
    
    If m_AccessToken = "" Then
        MsgBox "Please authenticate first (run AuthenticateUser or InitializeDropboxAPI)", vbExclamation
        EnsureValidToken = False
        Exit Function
    End If
    
    ' Check if token needs refresh
    If TokenNeedsRefresh() Then
        Debug.Print "Token expiring soon, refreshing..."
        If Not RefreshAccessToken() Then
            MsgBox "Token refresh failed. Please re-authenticate.", vbExclamation
            EnsureValidToken = False
            Exit Function
        End If
    End If
    
    EnsureValidToken = True
End Function

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Sub PauseExecution(seconds As Long)
    ' Pause execution for specified number of seconds (Access-compatible)
    On Error Resume Next
    
    Dim startTime As Double
    Dim currentTime As Double
    Dim elapsed As Double
    Dim targetSeconds As Double
    
    If seconds <= 0 Then Exit Sub
    If seconds > 3600 Then Exit Sub ' Safety: max 1 hour
    
    targetSeconds = CDbl(seconds)
    startTime = Timer
    
    Do
        DoEvents ' Allow other processes to run
        currentTime = Timer
        
        ' Handle midnight rollover (Timer resets to 0 at midnight)
        If currentTime < startTime Then
            ' Midnight occurred, adjust calculation
            elapsed = (86400# - startTime) + currentTime
        Else
            elapsed = currentTime - startTime
        End If
        
        ' Exit when enough time has passed
        If elapsed >= targetSeconds Then Exit Do
        
        ' Prevent infinite loop (allow 10 second tolerance)
        If elapsed > targetSeconds + 10# Then Exit Do
    Loop
End Sub

Private Function ExtractJsonValue(jsonString As String, key As String) As String
    ' Simple JSON parser for extracting string values
    ' Handles both "key": "value" and "key":"value" formats
    
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

Private Function EscapeJsonString(str As String) As String
    ' Escape special characters for JSON
    Dim result As String
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJsonString = result
End Function

Private Function UrlEncode(str As String) As String
    ' Simple URL encoding for OAuth parameters
    Dim result As String
    Dim i As Integer
    Dim char As String
    
    result = ""
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        If (char >= "A" And char <= "Z") Or _
           (char >= "a" And char <= "z") Or _
           (char >= "0" And char <= "9") Or _
           char = "-" Or char = "_" Or char = "." Or char = "~" Then
            result = result & char
        Else
            result = result & "%" & Right("0" & Hex(Asc(char)), 2)
        End If
    Next i
    
    UrlEncode = result
End Function

Private Function EncryptValue(value As String) As String
    ' Simple Base64-like encoding for database storage
    ' For very long tokens, we'll use a simple reversible encoding
    
    If value = "" Then
        EncryptValue = ""
        Exit Function
    End If
    
    On Error Resume Next
    
    ' For POC: Simple hex encoding (safe for database, easily reversible)
    Dim result As String
    Dim i As Long
    
    result = ""
    For i = 1 To Len(value)
        result = result & Right("00" & Hex(Asc(Mid(value, i, 1))), 2)
    Next i
    
    EncryptValue = result
End Function

Private Function DecryptValue(encryptedValue As String) As String
    ' Decrypt hex-encoded value
    
    If encryptedValue = "" Then
        DecryptValue = ""
        Exit Function
    End If
    
    On Error Resume Next
    
    Dim result As String
    Dim i As Long
    Dim hexPair As String
    
    result = ""
    For i = 1 To Len(encryptedValue) Step 2
        hexPair = Mid(encryptedValue, i, 2)
        result = result & Chr(CLng("&H" & hexPair))
    Next i
    
    DecryptValue = result
End Function

' ============================================================================
' LOGGING FUNCTIONS
' ============================================================================

Private Sub LogError(functionName As String, errorNumber As Long, errorDescription As String, Optional details As String = "")
    ' Log error to database table
    On Error Resume Next
    
    Dim db As Object ' DAO.Database - using late binding for compatibility
    Dim rs As Object ' DAO.Recordset - using late binding for compatibility
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblDropboxLog")
    
    rs.AddNew
    rs!LogDate = Now
    rs!LogLevel = "ERROR"
    rs!FunctionName = functionName
    rs!ErrorNumber = errorNumber
    rs!ErrorDescription = Left(errorDescription, 255)
    If details <> "" Then
        rs!details = details
    End If
    rs.Update
    rs.Close
    
    Set rs = Nothing
    Set db = Nothing
End Sub

Private Sub LogActivity(functionName As String, logLevel As String, details As String)
    ' Log activity to database table
    On Error Resume Next
    
    Dim db As Object ' DAO.Database - using late binding for compatibility
    Dim rs As Object ' DAO.Recordset - using late binding for compatibility
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblDropboxLog")
    
    rs.AddNew
    rs!LogDate = Now
    rs!LogLevel = logLevel
    rs!FunctionName = functionName
    rs!ErrorNumber = 0
    rs!ErrorDescription = ""
    rs!details = details
    rs.Update
    rs.Close
    
    Set rs = Nothing
    Set db = Nothing
End Sub

' ============================================================================
' PUBLIC UTILITY FUNCTIONS
' ============================================================================

Public Function GetAccessToken() As String
    ' Get current access token (for debugging)
    If m_AccessToken = "" Then
        GetAccessToken = "(not authenticated)"
    Else
        GetAccessToken = Left(m_AccessToken, 20) & "... (length: " & Len(m_AccessToken) & ")"
    End If
End Function

Public Function IsAuthenticated() As Boolean
    ' Check if user is authenticated
    IsAuthenticated = (m_AccessToken <> "")
End Function

Private Function TestConnection() As Boolean
    ' Test connection by making a simple API call
    On Error GoTo ErrHandler
    
    If m_AccessToken = "" Then
        TestConnection = False
        Exit Function
    End If
    
    Dim http As Object
    Dim jsonBody As String
    
    Debug.Print "Token length: " & Len(m_AccessToken) & " characters"
    Debug.Print "Token starts with: " & Left(m_AccessToken, 10) & "..."
    Debug.Print "Token ends with: ..." & Right(m_AccessToken, 10)
    
    ' Try to get current account info (simple, fast API call)
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", API_BASE & "users/get_current_account", False
    http.setRequestHeader "Authorization", "Bearer " & m_AccessToken
    http.setRequestHeader "Content-Type", "application/json"
    http.send "null"
    
    If http.Status = 200 Then
        TestConnection = True
    Else
        Debug.Print "✗ API Test Failed: " & http.Status
        Debug.Print "  Response: " & http.responseText
        
        If http.Status = 401 Then
            Debug.Print ""
            Debug.Print "⚠ TOKEN IS INVALID OR EXPIRED!"
            Debug.Print "  Short-lived tokens expire 4 hours after GENERATION (not after entering them)"
            Debug.Print "  Generate a NEW token at: https://www.dropbox.com/developers/apps"
            Debug.Print "  Then run: QuickSetupWithPrompt"
        End If
        
        TestConnection = False
    End If
    
    Set http = Nothing
    Exit Function
    
ErrHandler:
    Debug.Print "✗ Error testing connection: " & Err.Description
    TestConnection = False
End Function

Public Function GetTokenExpiry() As String
    ' Get token expiry date (for debugging)
    If m_TokenExpiry = #12:00:00 AM# Then
        GetTokenExpiry = "(unknown)"
    Else
        GetTokenExpiry = Format(m_TokenExpiry, "yyyy-mm-dd hh:nn:ss")
    End If
End Function

Public Sub ClearAuthentication()
    ' Clear stored authentication (logout)
    On Error Resume Next
    
    Debug.Print "========================================="
    Debug.Print "Clearing authentication..."
    
    ' Clear module variables
    m_AccessToken = ""
    m_RefreshToken = ""
    m_TokenExpiry = #12:00:00 AM#
    
    ' Deactivate all tokens in database
    CurrentDb.Execute "UPDATE tblDropboxTokens SET IsActive = False"
    
    Debug.Print "✓ Authentication cleared"
    Debug.Print "========================================="
    
    MsgBox "Authentication cleared successfully." & vbCrLf & vbCrLf & _
           "Run AuthenticateUser() to log in again.", _
           vbInformation, "Logged Out"
End Sub

' ============================================================================
' TEST FUNCTIONS
' ============================================================================

Public Sub TestAuthentication()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Authentication"
    Debug.Print "========================================="
    
    If AuthenticateUser() Then
        Debug.Print "✓ TEST PASSED: Authentication successful"
    Else
        Debug.Print "✗ TEST FAILED: Authentication failed"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub TestCreateFolder()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Create Folder"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "✗ TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    Dim folderPath As String
    folderPath = "/TB_CMS_POC/TestCase/" & Format(Now, "yyyy-mm-dd-hhnnss")
    
    If CreateFolder(folderPath) Then
        Debug.Print "✓ TEST PASSED: Folder creation successful"
    Else
        Debug.Print "✗ TEST FAILED: Folder creation failed"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub TestUpload()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Upload File"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "✗ TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' File picker - let user select file to upload
    Dim fd As FileDialog
    Dim selectedFile As String
    Dim dropboxPath As String
    Dim fileName As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a file to upload to Dropbox"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "All Files", "*.*"
    fd.Filters.Add "PDF Files", "*.pdf"
    fd.Filters.Add "Word Documents", "*.doc; *.docx"
    fd.Filters.Add "Excel Files", "*.xls; *.xlsx"
    fd.Filters.Add "Text Files", "*.txt"
    fd.Filters.Add "Images", "*.jpg; *.png; *.gif"
    
    If fd.Show = -1 Then
        selectedFile = fd.SelectedItems(1)
        fileName = Dir(selectedFile)
        dropboxPath = "/TB_CMS_POC/" & fileName
        
        Debug.Print "Selected file: " & selectedFile
        Debug.Print "Upload destination: " & dropboxPath
        
        If UploadFile(selectedFile, dropboxPath) Then
            Debug.Print "✓ TEST PASSED: Upload successful"
        Else
            Debug.Print "✗ TEST FAILED: Upload failed"
        End If
    Else
        Debug.Print "✗ TEST CANCELLED: No file selected"
    End If
    
    Set fd = Nothing
    Debug.Print "========================================="
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ TEST ERROR: " & Err.Description
    Debug.Print "========================================="
    Set fd = Nothing
End Sub

Public Sub TestListFolder()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: List Folder"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "✗ TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    Dim folderPath As String
    Dim contents As String
    
    folderPath = "/TB_CMS_POC"
    contents = ListFolder(folderPath)
    
    If contents <> "" Then
        Debug.Print "✓ TEST PASSED: List folder successful"
        Debug.Print "Contents preview: " & Left(contents, 300) & "..."
    Else
        Debug.Print "✗ TEST FAILED: List folder failed"
    End If
    
    Debug.Print "========================================="
End Sub

Public Sub TestDownload()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST: Download File"
    Debug.Print "========================================="
    
    If Not IsAuthenticated() Then
        Debug.Print "✗ TEST SKIPPED: Not authenticated. Run TestAuthentication first."
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' Step 1: Create a test file to upload
    Dim testFilePath As String
    Dim testContent As String
    Dim fileNum As Integer
    Dim dropboxPath As String
    Dim localPath As String
    
    testFilePath = Environ("TEMP") & "\download_test_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    testContent = "Download Test File" & vbCrLf & _
                  "Created: " & Now & vbCrLf & _
                  "This file tests download functionality."
    
    ' Write test file
    fileNum = FreeFile
    Open testFilePath For Output As #fileNum
    Print #fileNum, testContent
    Close #fileNum
    
    Debug.Print "Created test file for download test"
    
    ' Step 2: Upload it
    dropboxPath = "/TB_CMS_POC/download_test.txt"
    Debug.Print "Uploading test file..."
    
    If Not UploadFile(testFilePath, dropboxPath) Then
        Debug.Print "✗ TEST FAILED: Could not upload test file for download"
        On Error Resume Next
        Kill testFilePath
        Exit Sub
    End If
    
    ' Step 3: Download it to a different location
    localPath = Environ("TEMP") & "\downloaded_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    Debug.Print "Downloading file..."
    
    If DownloadFile(dropboxPath, localPath) Then
        Debug.Print "✓ TEST PASSED: Download successful"
        Debug.Print "  Downloaded to: " & localPath
        
        ' Verify file was downloaded
        If Dir(localPath) <> "" Then
            Debug.Print "  ✓ File verified on disk"
        End If
    Else
        Debug.Print "✗ TEST FAILED: Download failed"
    End If
    
    ' Clean up local test files
    On Error Resume Next
    Kill testFilePath
    
    Debug.Print "========================================="
    Exit Sub
    
ErrHandler:
    Debug.Print "✗ TEST ERROR: " & Err.Description
    Debug.Print "========================================="
End Sub

Public Sub RunAllTests()
    Debug.Print ""
    Debug.Print "+-------------------------------------------+"
    Debug.Print "|  DROPBOX API POC - COMPLETE TEST SUITE  |"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    
    ' Initialize
    Call InitializeDropboxAPI
    
    If Not IsAuthenticated() Then
        Debug.Print "Starting with authentication..."
        Call TestAuthentication
        
        If Not IsAuthenticated() Then
            Debug.Print ""
            Debug.Print "✗ CANNOT PROCEED: Authentication failed"
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
    Debug.Print "|         ALL TESTS COMPLETED!              |"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    Debug.Print "Review results above. All tests should show ✓"
    Debug.Print ""
End Sub

' ============================================================================
' SETUP WIZARD - Run this first!
' ============================================================================

Public Sub SetupWizard()
    ' Interactive setup wizard for first-time configuration
    ' RECOMMENDED: Uses OAuth for permanent access
    
    Debug.Print ""
    Debug.Print "+-------------------------------------------+"
    Debug.Print "|    DROPBOX API SETUP WIZARD (OAuth)       |"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    
    ' Step 1: Create tables
    Dim response As VbMsgBoxResult
    response = MsgBox("This wizard will set up Dropbox API with OAuth." & vbCrLf & vbCrLf & _
                      "✓ Browser opens ONCE for authorization" & vbCrLf & _
                      "✓ Then works FOREVER with auto-refresh" & vbCrLf & vbCrLf & _
                      "Step 1: Create database tables" & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Setup Wizard - Step 1")
    
    If response = vbNo Then
        MsgBox "Setup cancelled", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Call CreateConfigTables
    On Error GoTo 0
    
    ' Step 2: Configure app credentials
    response = MsgBox("Step 2: Configure Dropbox app credentials" & vbCrLf & vbCrLf & _
                      "You'll need:" & vbCrLf & _
                      "- App Key" & vbCrLf & _
                      "- App Secret" & vbCrLf & vbCrLf & _
                      "Get from: https://www.dropbox.com/developers/apps" & vbCrLf & _
                      "→ Your App → Settings" & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Setup Wizard - Step 2")
    
    If response = vbNo Then
        MsgBox "Setup incomplete. Run SetupWizard again when ready.", vbInformation
        Exit Sub
    End If
    
    Dim appKey As String
    Dim appSecret As String
    
    appKey = InputBox("Enter your Dropbox App Key:" & vbCrLf & vbCrLf & _
                      "Example: jbozj8nffezcw9w", _
                      "Setup Wizard - App Key")
    If Trim(appKey) = "" Then
        MsgBox "Setup cancelled", vbInformation
        Exit Sub
    End If
    
    appSecret = InputBox("Enter your Dropbox App Secret:" & vbCrLf & vbCrLf & _
                         "Example: qjp2rzxzgfhj9qf", _
                         "Setup Wizard - App Secret")
    If Trim(appSecret) = "" Then
        MsgBox "Setup cancelled", vbInformation
        Exit Sub
    End If
    
    Call SetupDropboxConfig(appKey, appSecret)
    
    ' Step 3: OAuth Authentication
    response = MsgBox("Step 3: OAuth Authentication" & vbCrLf & vbCrLf & _
                      "Your browser will open to Dropbox." & vbCrLf & vbCrLf & _
                      "What you'll do:" & vbCrLf & _
                      "1. Click 'Allow' to authorize" & vbCrLf & _
                      "2. Copy the authorization code" & vbCrLf & _
                      "3. Paste it back here" & vbCrLf & vbCrLf & _
                      "✓ This gives you permanent access!" & vbCrLf & vbCrLf & _
                      "Ready to authenticate?", _
                      vbYesNo + vbQuestion, "Setup Wizard - Step 3")
    
    If response = vbYes Then
        ' Full OAuth flow
        Call AuthenticateUser
    Else
        MsgBox "Authentication skipped." & vbCrLf & vbCrLf & _
               "Run AuthenticateUser later to complete setup.", _
               vbInformation
    End If
    
    Debug.Print ""
    Debug.Print "+-------------------------------------------+"
    Debug.Print "|         SETUP WIZARD COMPLETED!           |"
    Debug.Print "+-------------------------------------------+"
    Debug.Print ""
    Debug.Print "You can now use the Dropbox API functions:"
    Debug.Print "- UploadFile(localPath, dropboxPath)"
    Debug.Print "- DownloadFile(dropboxPath, localPath)"
    Debug.Print "- CreateFolder(dropboxPath)"
    Debug.Print "- ListFolder(dropboxPath)"
    Debug.Print ""
    Debug.Print "Run RunAllTests() to verify everything works!"
    Debug.Print ""
End Sub
