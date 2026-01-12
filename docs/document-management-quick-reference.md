# Document Management Quick Reference Guide

**Quick access guide for TB CMS Document Management features**

---

## User Interface Buttons (frmClientLedger)

### 📁 Folder Operations

| Button | Function | What It Does |
|--------|----------|--------------|
| **Create/Open Folder** | `cmdCreateFolder_Click()` | Opens the main case folder in Windows Explorer. Creates folder if it doesn't exist. |
| **Create Subfolder** | `cmdCreateFolderSub_Click()` | Opens a specific document type subfolder (e.g., Discovery, Correspondence). |
| **Open General Folder** | `cmdOpenDocumentFolderFull_Click()` | Opens the General documents folder. |
| **Open Correspondence** | `cmdOpenDocumentFolderCorrespondence_Click()` | Opens Correspondence: Letters and Emails folder. |
| **Open Discovery** | `cmdOpenDocumentFolderFinance_Click()` | Opens Discovery folder. |
| **Open Invoices** | `cmdOpenDocumentFolderInvoices_Click()` | Opens Client Invoices folder. |

### 📄 Document Operations

| Button | Function | What It Does |
|--------|----------|--------------|
| **Scan Document** | `cmdScan_Click()` | Main button to scan and save documents to case folder. |
| **Open Retainer** | `cmdOpenRetainer_Click()` | Opens the Retainer/Contract document. |
| **Open Initial Intake** | `cmdOpenInitialIntake_Click()` | Opens Initial Intake documents. |
| **Open Client ID** | `cmdOpenDocumentClientID_Click()` | Opens Client ID document. |
| **Open Closed Final** | `cmdOpenClosedFinal_Click()` | Opens Closed Final document. |

### 🗂️ Case Management

| Button | Function | What It Does |
|--------|----------|--------------|
| **Close Case** | `cmdCloseCase_Click()` | Closes case and optionally moves documents to _CLOSED folder and/or archives them. |
| **Reopen Case** | `cmdReopenCase_Click()` | Reopens closed case and optionally moves documents back to active area. |

---

## Document Types

### Standard Document Types

1. **General** - Main case documents (default)
2. **Init Intake, Notes, Documents** - Initial intake paperwork
3. **Client ID** - Client identification documents
4. **Retainer / Contract** - Retainer agreements
5. **Correspondence: Letters and Emails** - Communications
6. **Discovery** - Discovery documents
7. **Client Invoices** - Billing invoices
8. **Closed Final** - Final closing documents

### Special Document Types

- **Intake Documents** - Pre-case documents (for potential clients)
- **Closed File Scans** - Archived documents for permanently closed cases

---

## Common Workflows

### ✅ Scan and Save a Document

**Steps:**
1. Open case in `frmClientLedger`
2. Select document type from dropdown (e.g., "Retainer / Contract")
3. Click **Scan Document** button
4. In file picker, navigate to scanned file (usually in Scanner folder)
5. Select file and click Open
6. System suggests filename: `2023-0123-Smith_John-Retainer.pdf`
7. Confirm or modify filename
8. Click Save
9. ✅ Document saved and linked to case

**Result:** Document is copied to case folder and record saved to database.

---

### 📂 Open Case Folder

**Steps:**
1. Open case in `frmClientLedger`
2. Click **Create/Open Folder** button
3. If folder doesn't exist, click **Yes** when prompted to create
4. Windows Explorer opens to case folder
5. Browse, add, or organize files as needed

**Result:** Case folder opened in Windows Explorer for manual file management.

---

### 📝 Open a Specific Document

**Steps:**
1. Open case in `frmClientLedger`
2. Click appropriate button:
   - **Open Retainer** for retainer agreement
   - **Open Initial Intake** for intake documents
   - **Open Client ID** for client ID
   - **Open Closed Final** for closing documents
3. Document opens in default application (e.g., Adobe Reader for PDFs)

**Result:** Document opens for viewing/editing.

**Note:** If document doesn't exist, error message appears.

---

### 🗃️ Close a Case with Documents

**Steps:**
1. Open case in `frmClientLedger`
2. Verify AR Balance = $0 and Trust Balance = $0
3. Click **Close Case** button
4. System marks case as closed
5. **Prompt 1:** "Move to CLOSED FILE SCANS?"
   - **Yes** → Copies entire case folder to archive location
   - **No** → Skip archiving
6. **Prompt 2:** "Move to _CLOSED subfolder?"
   - **Yes** → Moves folder from active area to _CLOSED
   - **No** → Leave folder in active area

**Result:** Case closed, documents archived and/or moved to closed area.

---

### 🔄 Reopen a Closed Case

**Steps:**
1. Open closed case in `frmClientLedger`
2. Click **Reopen Case** button
3. **Prompt:** "Move back the Client folder?"
   - **Yes** → Moves folder from _CLOSED back to active area
   - **No** → Leave folder in _CLOSED
4. Case reopened

**Result:** Case status changed to Open, documents optionally moved back.

---

## Folder Structure

### Active Case Folder Structure

```
S:\Client Files\
└── 2023-Smith_John\               ← Case Folder
    ├── General\                   ← Default documents
    ├── ClientID\                  ← Client identification
    ├── Retainer\                  ← Retainer agreements
    ├── Correspondence\            ← Letters & emails
    ├── Discovery\                 ← Discovery documents
    ├── Invoices\                  ← Client invoices
    ├── ClosedFinal\               ← Closing documents
    └── Intake\                    ← Initial intake (if applicable)
```

### Closed Case Folder Structure

```
S:\Client Files\
└── _CLOSED\                       ← Closed cases subdirectory
    └── 2023-Smith_John\           ← Case Folder (moved from active)
        ├── General\
        ├── ClientID\
        ├── Retainer\
        ├── Correspondence\
        ├── Discovery\
        ├── Invoices\
        ├── ClosedFinal\
        └── Intake\
```

### Archive Structure

```
S:\Closed File Scans\              ← Permanent archive
└── 2023-Smith_John\               ← Copy of entire case folder
    └── [All subfolders copied]
```

---

## File Naming Convention

### Standard Format
```
{Year}-{CaseNumber}-{LastName}_{FirstName}-{DocumentType}.{extension}
```

### Examples
- `2023-0123-Smith_John-RetainerAgreement.pdf`
- `2023-0123-Smith_John-ClientID.pdf`
- `2023-0456-Jones_Mary-ClosingDocuments.docx`
- `2024-0789-Brown_Bob-Correspondence_2024-01-12.pdf`

---

## Scanner Workflow

### Setup
1. Physical document → Scanner
2. Scan to: `S:\Scanner\`
3. Filename: Auto-generated by scanner (e.g., `scan_20240112_001.pdf`)

### Import to Case
1. In TB CMS, open case
2. Select document type
3. Click **Scan Document**
4. Select file from `S:\Scanner\`
5. System copies to case folder with standardized name
6. Document linked to case in database

---

## Keyboard Shortcuts (if available)

*Note: Check if any keyboard shortcuts are configured in the form.*

Currently, buttons require mouse clicks. Consider adding:
- `Alt+F` → Open folder
- `Alt+S` → Scan document
- `Alt+O` → Open document

---

## Troubleshooting

### ❌ "Folder doesn't exist. Create it?"
**Cause:** Case folder not yet created  
**Solution:** Click **Yes** to create folder automatically

---

### ❌ "Failed to open folder"
**Cause:** Invalid folder path or insufficient permissions  
**Solution:** 
1. Verify case information is correct
2. Check network drive is connected (e.g., S:\ drive)
3. Verify Windows permissions for folder

---

### ❌ "Document is not found"
**Cause:** Document file missing or moved  
**Solution:**
1. Open case folder manually
2. Verify file exists
3. If file moved, re-scan/re-save document
4. If database reference broken, may need IT support

---

### ❌ "Case is closed!"
**Cause:** Attempting to scan documents to closed case  
**Solution:** Reopen case first, then scan documents

---

### ❌ "Unable to delete original folder after copy"
**Cause:** Files are locked or in use  
**Solution:** 
1. Close any open documents from that folder
2. Manually delete folder after closing files
3. Contact IT if issue persists

---

### ❌ "Please select a case before proceeding"
**Cause:** No case selected in form  
**Solution:** Select case from dropdown or search before clicking button

---

## Best Practices

### ✅ DO:
- ✅ Select correct document type before scanning
- ✅ Use descriptive filenames when saving
- ✅ Verify document saved successfully (check success message)
- ✅ Close case only when balances are $0
- ✅ Archive important cases to Closed File Scans

### ❌ DON'T:
- ❌ Don't manually move files outside the system
- ❌ Don't delete case folders manually (use system)
- ❌ Don't scan documents to closed cases
- ❌ Don't modify folder structure (breaks database links)
- ❌ Don't bypass Save As dialog (system needs to track files)

---

## Configuration

### System Settings

**Location:** `tblDocumentRootDirectory` table in database

| Setting | Example Value | Description |
|---------|--------------|-------------|
| `DocumentRootDirectory` | `S:\Client Files\` | Root path for all case documents |
| `ScannerDirectory` | `S:\Scanner\` | Temporary folder for scanned files |
| `ClosedFileScanDirectory` | `S:\Closed File Scans\` | Archive location for closed cases |

**To modify:** Contact database administrator (requires SQL Server access)

---

## Permissions Required

### File System Permissions
- **Read/Write** on `DocumentRootDirectory` (e.g., `S:\Client Files\`)
- **Read** on `ScannerDirectory` (e.g., `S:\Scanner\`)
- **Write** on `ClosedFileScanDirectory` (e.g., `S:\Closed File Scans\`)

### Database Permissions
- **Execute** permissions on stored procedures:
  - `spGetDocumentFolderName`
  - `spSaveCaseDocument`
  - `spGetCaseDocument`
  - (and others - see full list in technical docs)

### Application Permissions
- Access to `frmClientLedger` form
- Case read/write permissions

---

## Related Forms

### Forms Using Document Management

1. **frmClientLedger** - Main form (most features)
2. **frmPersInjProvider** - Opens medical documents folder
3. **Time Keeping** - Opens invoice folder
4. **frmInvoiceSent** - Opens invoice folder

---

## Support & Help

### For Technical Issues:
- Check network drive connectivity
- Verify file permissions
- Contact IT support for server issues

### For Workflow Questions:
- Refer to this guide
- See `docs/document-management-analysis.md` for detailed technical documentation
- Contact system administrator for custom workflows

---

## Quick Tips

💡 **Tip 1:** Use "General" document type for most documents. Create subfolders only when needed for organization.

💡 **Tip 2:** Always verify the success message after scanning documents. If no message appears, document may not have saved.

💡 **Tip 3:** Before closing a case, decide whether to archive (Closed File Scans) or just move to _CLOSED. Archive takes extra disk space but provides backup.

💡 **Tip 4:** When reopening a case, remember to move documents back from _CLOSED if you moved them there.

💡 **Tip 5:** If you can't find a document, try opening the case folder directly and browsing subfolders.

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-01-12 | Initial quick reference guide created |

---

**For detailed technical documentation, see:**
- `docs/document-management-analysis.md` - Full technical analysis
- `docs/document-management-summary.md` - Executive summary
