# Dropbox POC - Quick Start Guide

**Goal**: Get a working prototype running in 1-2 weeks  
**Prerequisites**: MS Access, Dropbox account, Internet

---

## Day 1: Setup (4 hours)

### Step 1: Register Dropbox App (30 minutes)

1. Go to https://www.dropbox.com/developers/apps
2. Click **"Create app"**
3. Choose:
   - **API**: Scoped access
   - **Access**: Full Dropbox
   - **Name**: `TB-CMS-POC`
4. Click **"Create app"**
5. In **Settings** tab:
   - Copy **App key** (save it!)
   - Copy **App secret** (save it!)
   - Under **OAuth 2**, add redirect URI: `http://localhost`
6. In **Permissions** tab, enable:
   - ✅ `files.metadata.write`
   - ✅ `files.metadata.read`
   - ✅ `files.content.write`
   - ✅ `files.content.read`
   - Click **Submit**

**You now have:**
- App Key: `abc123...`
- App Secret: `xyz789...`

---

### Step 2: Create Test Dropbox Folder (15 minutes)

1. Log into Dropbox web
2. Create folder: `/TB_CMS_POC`
3. Create subfolder: `/TB_CMS_POC/TestCase`
4. Upload a test PDF manually

---

### Step 3: Set Up MS Access (30 minutes)

1. Open MS Access
2. Create new blank database: `DropboxPOC.accdb`
3. Go to **File → Options → Trust Center → Trust Center Settings**
4. Click **Macro Settings**
5. Enable **"Trust access to the VBA project object model"**
6. Click OK

**Add References:**
1. Press `Alt+F11` (open VBA editor)
2. Go to **Tools → References**
3. Check:
   - ✅ Microsoft XML, v6.0
   - ✅ Microsoft Scripting Runtime
   - ✅ Microsoft Office 16.0 Object Library
4. Click OK

---

### Step 4: Create API Module (2 hours)

1. In VBA Editor, click **Insert → Module**
2. Rename module to `DropboxAPI_POC`
3. Copy the full code from `dropbox-poc-plan.md` (Phase 2, Step 2.1)
4. **IMPORTANT**: Replace these lines at the top:
   ```vba
   Private Const DROPBOX_APP_KEY As String = "YOUR_APP_KEY_HERE"
   Private Const DROPBOX_APP_SECRET As String = "YOUR_APP_SECRET_HERE"
   ```
   With your actual App Key and App Secret from Step 1

5. Save (`Ctrl+S`)

---

### Step 5: Create Temp Folder (5 minutes)

1. Open Windows Explorer
2. Create folder: `C:\Temp`
3. This is where downloaded files will go

---

## Day 2: Test Authentication (30 minutes)

### Test OAuth 2.0 Flow

1. In VBA Editor, press `Ctrl+G` to open Immediate Window
2. Type: `DropboxAPI_POC.TestAuthentication`
3. Press Enter
4. Browser will open to Dropbox authorization page
5. Click **"Allow"**
6. Copy the authorization code shown
7. Paste into Access prompt
8. Click OK

**Expected**: Message box "Authentication successful!"

**If it fails:**
- Verify App Key and App Secret are correct
- Check Permissions are enabled in Dropbox app settings
- Check redirect URI is exactly `http://localhost`

---

## Day 3-4: Test File Operations

### Test Create Folder

**Immediate Window:**
```vba
DropboxAPI_POC.TestCreateFolder
```

**Expected**: Folder `/TB_CMS_POC/TestCase/2023-Smith_John/General` created

---

### Test Upload

**Immediate Window:**
```vba
DropboxAPI_POC.TestUpload
```

1. File picker will open
2. Select a PDF file (< 5MB recommended)
3. Wait for upload

**Expected**: Message "File uploaded successfully!" with time

---

### Test List Folder

**Immediate Window:**
```vba
DropboxAPI_POC.TestListFolder
```

**Expected**: JSON response with folder contents shown in Immediate Window

---

### Test Download

**Immediate Window:**
```vba
DropboxAPI_POC.TestDownload
```

**Expected**: 
- File downloaded to `C:\Temp\downloaded_test.pdf`
- File opens automatically

---

### Run All Tests

**Immediate Window:**
```vba
DropboxAPI_POC.RunAllTests
```

**Expected**: All tests run automatically with results in Immediate Window

---

## Day 5: Build Test Form (Optional but Recommended)

### Create Form

1. In Access, click **Create → Form Design**
2. Add these controls:
   - **Button**: `btnAuthenticate` - Caption: "Authenticate"
   - **Button**: `btnUpload` - Caption: "Upload File"
   - **Button**: `btnDownload` - Caption: "Download File"
   - **Button**: `btnCreateFolder` - Caption: "Create Folder"
   - **Button**: `btnListFolder` - Caption: "List Folder"
   - **Button**: `btnRunAllTests` - Caption: "Run All Tests"
   - **TextBox**: `txtStatus` - Label: "Status:"
   - **TextBox**: `txtLog` - Multi-line, scrollbars, large

3. Right-click form → **View Code**
4. Copy form code from `dropbox-poc-plan.md` (Phase 3, Step 3.1)
5. Save form as `frmDropboxPOC`

### Test Form

1. Open form
2. Click each button
3. Verify all functions work

---

## Troubleshooting

### Problem: "Please authenticate first"
**Solution**: Click "Authenticate" button first before other operations

---

### Problem: "Authentication failed"
**Solutions:**
- Verify App Key and Secret are correct (no extra spaces)
- Check Dropbox app permissions are enabled
- Try re-creating the Dropbox app
- Check internet connection

---

### Problem: "Upload failed: 401"
**Solution**: Token expired. Click "Authenticate" again

---

### Problem: "Upload failed: 409"
**Solution**: File already exists. Either:
- Delete file from Dropbox first
- Or change the upload path in code

---

### Problem: "Download failed: 404"
**Solution**: File doesn't exist at that path. Verify:
- Path is correct (case-sensitive!)
- File was uploaded successfully
- Path starts with `/` 

---

### Problem: Code runs but nothing happens
**Solutions:**
- Check Immediate Window (`Ctrl+G`) for Debug.Print output
- Add breakpoints (`F9`) to debug
- Verify m_AccessToken is not empty after auth

---

## Performance Benchmarks

Record your results:

| Operation | File Size | Your Time | Target |
|-----------|-----------|-----------|--------|
| Upload | 100 KB | ______ s | < 1s |
| Upload | 1 MB | ______ s | < 3s |
| Upload | 5 MB | ______ s | < 15s |
| Download | 100 KB | ______ s | < 1s |
| Download | 1 MB | ______ s | < 3s |
| Download | 5 MB | ______ s | < 15s |
| Create Folder | N/A | ______ s | < 1s |
| List Folder | 10 files | ______ s | < 2s |

**If times are much slower:**
- Check your internet speed
- Try smaller files first
- Network congestion?

---

## What's Next?

### If POC is Successful:
1. ✅ Document what worked well
2. ✅ Note any challenges
3. ✅ Demo to stakeholders
4. ✅ Get approval for full implementation
5. ✅ Start Phase 1 of full development plan

### If POC Has Issues:
1. Document specific problems
2. Research solutions
3. Refine approach
4. Run targeted tests
5. Re-assess feasibility

---

## Quick Reference

### Essential Functions

```vba
' Authenticate (do this first!)
DropboxAPI_POC.AuthenticateUser()

' Upload a file
DropboxAPI_POC.UploadFile("C:\local\file.pdf", "/dropbox/path/file.pdf")

' Download a file
DropboxAPI_POC.DownloadFile("/dropbox/path/file.pdf", "C:\local\file.pdf")

' Create folder
DropboxAPI_POC.CreateFolder("/dropbox/path/folder")

' List folder
contents = DropboxAPI_POC.ListFolder("/dropbox/path")

' Check if authenticated
isAuth = DropboxAPI_POC.IsAuthenticated()

' Get current token
token = DropboxAPI_POC.GetAccessToken()
```

---

## Common Paths

```vba
' Test folder
"/TB_CMS_POC"

' Test case folder
"/TB_CMS_POC/TestCase/2023-Smith_John/General"

' Test upload
"/TB_CMS_POC/test_upload.pdf"
```

---

## Support

### Official Resources
- **Dropbox API Docs**: https://www.dropbox.com/developers/documentation/http/overview
- **OAuth Guide**: https://www.dropbox.com/developers/reference/oauth-guide
- **API Explorer**: https://dropbox.github.io/dropbox-api-v2-explorer/

### Debug Tips
1. Always check Immediate Window (`Ctrl+G`) for Debug.Print output
2. Use breakpoints (`F9`) to step through code
3. Check http.Status for error codes (200 = success, 401 = auth error, 409 = conflict)
4. Verify Dropbox web interface to see if files actually uploaded

---

## Success Checklist

By end of POC, you should have:

- [ ] Successfully authenticated with Dropbox
- [ ] Uploaded at least 3 files (different sizes)
- [ ] Downloaded and opened files
- [ ] Created nested folder structure
- [ ] Listed folder contents
- [ ] Documented performance metrics
- [ ] Identified any issues or limitations
- [ ] Demo-ready proof of concept

---

## Time Estimate

| Task | Estimated Time | Your Actual Time |
|------|---------------|------------------|
| Setup (Day 1) | 4 hours | _______ |
| Test Auth (Day 2) | 30 minutes | _______ |
| Test Operations (Day 3-4) | 4 hours | _______ |
| Build Form (Day 5) | 3 hours | _______ |
| Documentation | 2 hours | _______ |
| **Total** | **~14 hours** | _______ |

**Spread over 5-7 calendar days** (allowing for interruptions)

---

## Ready to Start?

1. ✅ Read this guide completely
2. ✅ Have Dropbox account ready
3. ✅ Have MS Access installed
4. ✅ Have 2-3 hours blocked for initial setup
5. ✅ Have test PDF files ready (various sizes)

**Then jump to Day 1, Step 1 and start building!**

---

**Questions? Issues? Stuck?**

1. Check the Troubleshooting section above
2. Review the full POC plan in `dropbox-poc-plan.md`
3. Search Dropbox API documentation
4. Check your Debug.Print output in Immediate Window

**Good luck! You've got this! 🚀**
