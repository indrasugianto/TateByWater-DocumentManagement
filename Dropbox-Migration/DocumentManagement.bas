Attribute VB_Name = "DocumentManagement"
Option Compare Database
Option Explicit
Dim foo

' =============================================================================
' Phase 4 rewire status (last updated 2026-05-15):
'
'   Phase 4a — Read-flow rewire (DONE):
'     OpenDocumentFile     -> DropboxService.OpenDocument
'                             (download to %TEMP%\TBCMS\ + native-app launch)
'     OpenDocumentFolder   -> Dropbox web URL (Application.FollowHyperlink)
'
'   Phase 4b — Config layer + local-synced-root routing (DONE):
'     OpenDocumentFolder   -> Windows Explorer against the local-synced
'                             folder when the desktop client is signed in,
'                             with fallback to the Dropbox web URL.
'                             Eliminates the "Unsupported path provided"
'                             warning Dropbox emits for team-namespace
'                             deep-links via /home/ or /work/.
'     GetDocumentRootFolder -> read tblDropboxRootConfig.TeamRootPath +
'                              DropboxService.DropboxPathToLocalPath
'                              (legacy: tblDocumentRootDirectory).
'                              Currently has zero callers in the project
'                              (kept for contract stability).
'     GetScannerFolder     -> read tblDropboxRootConfig.ScannerDirectory +
'                             DropboxService.DropboxPathToLocalPath.
'                             Two callers: Intakes.cmdScan_Click + the
'                             frmClientLedger scan flow. Both consume the
'                             local Windows path as Office.FileDialog
'                             InitialFileName.
'
'   Phase 4c — Stored procedure rewrites (DONE in Dropbox-Migration-SQL-
'              Install.sql Section 8). Live on awsql2022dev/TateByWater.
'
'   Phase 4d — Write-flow rewires (DONE):
'     SaveScannedFileAs    -> override-with-confirmation Save-As prompt +
'                             size-routed DropboxService.UploadFile /
'                             UploadLargeFile + SaveCaseDocument (with G13
'                             token guard) + orphan-file compensation
'                             policy (compensating DeleteFile on SP failure;
'                             tblDropboxOrphanQueue INSERT via private
'                             helper QueueOrphanFile if delete also fails).
'     MoveDocumentByCaseStatus -> DropboxService.MoveFile + new G2
'                             spMoveDocumentFolder (@OldFolderPath /
'                             @NewFolderPath) + three-branch SP outcome
'                             handling (both>0 commit, exactly-one-zero
'                             warn+commit, throw → reverse-move
'                             compensation + critical-state surface if
'                             reverse also fails).
'     CopyDocumentToClosedFileScan -> DropboxService.CopyFile against
'                             <ClosedFileScanRoot>/<basename(source)>.
'
'   Test-environment behavior of every 4d write call: hits
'   DropboxService.GuardWritesEnabled and raises immediately
'   (ALLOW_DROPBOX_WRITES = False). The error propagates back to the
'   DocumentManagement function, which surfaces a generic user-facing
'   message and returns False. No real Dropbox traffic, no SQL writes.
'
'   Rollback safety: legacy (pre-4a / pre-4b) bodies of every rewired
'   function are preserved as commented-out blocks immediately above each
'   active function. Swap procedure documented above the first LEGACY block.
' =============================================================================


Public Function GetDocumentFileName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetDocumentFileName "
    sql = sql & "@CaseID = " & CaseID
    sql = sql & ",@DocumentType = " & pcaAddQuotes(DocumentType)

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = pcaConvertNulls(rs("FileName"), "")
    Else
        rv = ""
    End If
Exit_Handler:
    GetDocumentFileName = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


Public Function GetDocumentFolderName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetDocumentFolderName "
    sql = sql & "@CaseID = " & CaseID
    sql = sql & ",@DocumentType = " & pcaAddQuotes(DocumentType)

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = pcaConvertNulls(rs("DocumentFolder"), "")
    Else
        rv = ""
    End If
Exit_Handler:
    GetDocumentFolderName = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


Public Function GetIntakeFolderName() As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetIntakeFolderName "

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = pcaConvertNulls(rs("DocumentFolder"), "")
    Else
        rv = ""
    End If
Exit_Handler:
    GetIntakeFolderName = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


Public Function GetClosedDocumentFolderName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetClosedDocumentFolderName "
    sql = sql & "@CaseID = " & CaseID
    sql = sql & ",@DocumentType = " & pcaAddQuotes(DocumentType)

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = pcaConvertNulls(rs("DocumentFolder"), "")
    Else
        rv = ""
    End If
Exit_Handler:
    GetClosedDocumentFolderName = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


' ============================================================================
' LEGACY (pre-Phase 4b) — GetDocumentRootFolder
' ----------------------------------------------------------------------------
' Preserved for rollback. Same swap procedure as the LEGACY OpenDocumentFolder
' block earlier in this file (comment-out active + uncomment LEGACY + re-import).
' ============================================================================
' Public Function GetDocumentRootFolder() As String
' On Error GoTo Err_Handler
' Dim rv As String
' Dim cn As ADODB.Connection
' Dim rs As ADODB.Recordset
' Dim sql As String
'     Set cn = New ADODB.Connection
'     cn.Open PcaGetConnnectionString
'     sql = "SELECT DocumentRootDirectory FROM tblDocumentRootDirectory"
'     Set rs = cn.Execute(sql)
'     If Not rs.EOF() Then
'         rv = rs("DocumentRootDirectory")
'     Else
'         rv = ""
'     End If
' Exit_Handler:
'     GetDocumentRootFolder = rv
'     Exit Function
' Err_Handler:
'     rv = ""
'     foo = pcaStdErrMsg(Err, Error)
'     Resume Exit_Handler
' End Function
' ============================================================================


' --- Phase 4b rewire --------------------------------------------------------
' Reads TeamRootPath from tblDropboxRootConfig (e.g., /Company/COMMON) and
' converts to the local-synced Windows path (e.g.,
' C:\Users\<u>\Tate Bywater Dropbox\<Name>\Company\COMMON) so the result is
' compatible with the legacy contract (a Windows path callers can pass to
' Office.FileDialog or Dir).
'
' Returns "" when the desktop client isn't installed/signed in
' (m_LocalSyncedRoot unresolved). Callers must tolerate "" (the existing
' implementation also returned "" if tblDocumentRootDirectory was empty,
' so this is contract-compatible).
'
' This function is currently unreferenced by other VBA in the project
' (Phase 4b survey: zero callers). Updated for consistency.
Public Function GetDocumentRootFolder() As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim dropboxPath As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = "SELECT TeamRootPath FROM dbo.tblDropboxRootConfig WHERE ConfigID = 1"

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        dropboxPath = pcaConvertNulls(rs("TeamRootPath"), "")
    End If

    If LenB(dropboxPath) > 0 Then
        rv = DropboxService.DropboxPathToLocalPath(dropboxPath)
    Else
        rv = ""
    End If
Exit_Handler:
    GetDocumentRootFolder = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


' ============================================================================
' LEGACY (pre-Phase 4b) — GetScannerFolder
' ============================================================================
' Public Function GetScannerFolder() As String
' On Error GoTo Err_Handler
' Dim rv As String
' Dim cn As ADODB.Connection
' Dim rs As ADODB.Recordset
' Dim sql As String
'     Set cn = New ADODB.Connection
'     cn.Open PcaGetConnnectionString
'     sql = "SELECT ScannerDirectory FROM tblDocumentRootDirectory"
'     Set rs = cn.Execute(sql)
'     If Not rs.EOF() Then
'         rv = rs("ScannerDirectory")
'     Else
'         rv = ""
'     End If
' Exit_Handler:
'     GetScannerFolder = rv
'     Exit Function
' Err_Handler:
'     rv = ""
'     foo = pcaStdErrMsg(Err, Error)
'     Resume Exit_Handler
' End Function
' ============================================================================


' --- Phase 4b rewire --------------------------------------------------------
' Reads ScannerDirectory from tblDropboxRootConfig (e.g.,
' /Company/COMMON/_SCANNER) and converts to the local-synced Windows path
' so callers — Intakes.cmdScan_Click + frmClientLedger's scan flow — can
' pass it to Office.FileDialog as the InitialFileName.
'
' Returns "" when the desktop client isn't installed/signed in. Callers
' fall back to whatever default Office.FileDialog picks (typically the
' user's last-used folder).
Public Function GetScannerFolder() As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim dropboxPath As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = "SELECT ScannerDirectory FROM dbo.tblDropboxRootConfig WHERE ConfigID = 1"

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        dropboxPath = pcaConvertNulls(rs("ScannerDirectory"), "")
    End If

    If LenB(dropboxPath) > 0 Then
        rv = DropboxService.DropboxPathToLocalPath(dropboxPath)
    Else
        rv = ""
    End If
Exit_Handler:
    GetScannerFolder = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


Public Function GetClosedFileScanFolderName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetClosedFileScanFolderName "
    sql = sql & "@CaseID = " & CaseID
    sql = sql & ",@DocumentType = " & pcaAddQuotes(DocumentType)

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = pcaConvertNulls(rs("DocumentFolder"), "")
    Else
        rv = ""
    End If
Exit_Handler:
    GetClosedFileScanFolderName = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


Public Function GetAllInvoicesFolderName(ByVal CaseID As Long) As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetAllInvoicesFolderName "
    sql = sql & "@CaseID = " & CaseID

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = pcaConvertNulls(rs("DocumentFolder"), "")
    Else
        rv = ""
    End If
Exit_Handler:
    GetAllInvoicesFolderName = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


' Local-filesystem helper. Obsolete in the Dropbox world (Dropbox auto-creates
' folders on upload, and folder-existence checks go through DropboxService.
' GetMetadata). Kept for now because Phase 4d still has legacy callers in
' the unchanged write-flow functions below; will be removed when 4d lands.
Function FolderExistsCreate(DirectoryPath As String, CreateIfNot As Boolean) As Boolean
    On Error GoTo Err_Handler
    Dim rv As Boolean
    Dim elm As Variant
    Dim strCheckPath As String
    If Right(DirectoryPath, 1) <> "\" Then
        DirectoryPath = DirectoryPath & "\"
    End If

    If Dir(DirectoryPath, vbDirectory) <> "" Then
        rv = True
    Else
        If CreateIfNot Then
            strCheckPath = ""
            For Each elm In Split(DirectoryPath, "\")
                strCheckPath = strCheckPath & elm & "\"
                If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
            Next
            rv = True
        Else
            rv = False
        End If
    End If
Exit_Handler:
    FolderExistsCreate = rv
    Exit Function
Err_Handler:
    rv = False
    Resume Exit_Handler
End Function


' Legacy local file picker — opens a dialog rooted at StartingFolder, then
' FollowHyperlinks the picked file. Used by the legacy OpenDocumentFolder
' which is replaced in 4a. Kept for any other callers; pending audit.
Public Function OpenFileDialog(ByVal DialogBoxTitle As String, ByVal StartingFolder As String, ByVal FileExtention As String)
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant

   Set fDialog = Application.FileDialog(msoFileDialogOpen)

   With fDialog
      .AllowMultiSelect = False
      .Title = DialogBoxTitle
      .InitialFileName = StartingFolder
      .Filters.Clear
      .Filters.Add "All Files", "*.*"

      If .show = True Then
         For Each varFile In .SelectedItems
            Application.FollowHyperlink varFile
         Next
      End If
   End With
End Function


' Local file picker — returns the selected local path. Still useful in the
' Dropbox world: ingest workflows (scan-save, intake) let the user pick a
' file from their Dropbox-desktop-synced folder, then call
' DropboxService.LocalPathToDropboxPath (added in 4b) to convert the
' Windows path to a Dropbox API path before upload.
Public Function SelectFileDialog(ByVal DialogBoxTitle As String, ByVal StartingFolder As String, ByVal FileExtention As String) As String
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   Dim rv As String

   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

   With fDialog
      .AllowMultiSelect = False
      .Title = DialogBoxTitle
      .InitialFileName = StartingFolder
      .Filters.Clear
      .Filters.Add "All Files", "*.*"

      If .show = True Then
         For Each varFile In .SelectedItems
            rv = varFile
         Next
      Else
        rv = ""
      End If
   End With
   SelectFileDialog = rv
End Function


' ============================================================================
' LEGACY (pre-Phase 4d) — SaveScannedFileAs
' ----------------------------------------------------------------------------
' Preserved for rollback. Used S:\ FileCopy directly + spSaveCaseDocument.
' ============================================================================
' Public Function SaveScannedFileAs(ByVal CaseID As Integer, ByVal DocumentType As String, ByVal SourceFileName As String, ByVal CaseStatus As String) As Boolean
' On Error GoTo Err_Handler
' Dim rv As Boolean
' Dim FolderName As String
' Dim DestinationFileName As String
' Dim fDialog As Office.FileDialog
' Dim varFile As Variant
'     If CaseStatus = "Closed" Then
'         FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
'     Else
'         FolderName = GetDocumentFolderName(CaseID, DocumentType)
'     End If
'     DestinationFileName = GetDocumentFileName(CaseID, DocumentType)
'     DestinationFileName = DestinationFileName & "." & Right(SourceFileName, Len(SourceFileName) - InStrRev(SourceFileName, "."))
'     If FolderExistsCreate(FolderName, True) Then
'         Set fDialog = Application.FileDialog(msoFileDialogSaveAs)
'         With fDialog
'             .AllowMultiSelect = False
'             .InitialFileName = FolderName
'             .InitialFileName = FolderName & DestinationFileName
'             If .show = True Then
'                 For Each varFile In .SelectedItems
'                     FileCopy SourceFileName, varFile
'                     If DocumentType = "Closed Final" Then
'                         If MsgBox("Do you want to save the file in Closed File Scans directory?", vbYesNo, "TB CMS") = vbYes Then
'                             FolderName = GetClosedFileScanFolderName(CaseID, "General")
'                             If FolderExistsCreate(FolderName, True) Then
'                                 FileCopy SourceFileName, FolderName & DestinationFileName
'                             End If
'                         End If
'                     End If
'                     If Not SaveCaseDocument(CaseID, DocumentType, varFile) Then
'                         MsgBox "Fail to save case document record...", , "TB CMS"
'                     End If
'                 Next
'             End If
'         End With
'     End If
'     rv = True
' Exit_Handler:
'     SaveScannedFileAs = rv
'     Exit Function
' Err_Handler:
'     rv = False
'     Resume Exit_Handler
' End Function
' ============================================================================


' --- Phase 4d rewire --------------------------------------------------------
' Scan-save workflow (Phase 5 step 2 in the plan). External-source ingest:
' SourceFileName is a local Windows path the user already picked (typically
' from the Dropbox-desktop-synced scanner folder).
'
' Sequence:
'   1. Resolve the Dropbox destination folder (Closed vs Open case path)
'      and the SP-generated filename. Append source's extension.
'   2. Override-with-confirmation: show a SaveAs dialog rooted in
'      %TEMP%\TBCMS\ pre-filled with the SP-generated filename. User can
'      edit; if changed, confirm before proceeding. (We use the temp dir
'      because the real destination lives on Dropbox — Office.FileDialog
'      can't navigate Dropbox API paths.)
'   3. Size-route to UploadFile (<= 150 MB) or UploadLargeFile (> 150 MB).
'      Upload reads directly from SourceFileName — no local copy.
'   4. On upload success: register via SaveCaseDocument (which now includes
'      the G13 token-validation guard).
'   5. Closed Final special-case: optionally also copy to ClosedFileScan
'      via DropboxService.CopyFile.
'   6. Orphan compensation policy: if SaveCaseDocument fails after a
'      successful upload, attempt to delete the upload. If the delete also
'      fails, queue the orphan to tblDropboxOrphanQueue for IT-admin drain.
'
' Test-environment behavior: every DropboxService write call hits
' GuardWritesEnabled and raises vbObjectError + 6001 (ALLOW_DROPBOX_WRITES =
' False). The error propagates here; we surface a user-facing message and
' return False. No SQL record is written, no real upload happens.
Public Function SaveScannedFileAs(ByVal CaseID As Integer, _
                                   ByVal DocumentType As String, _
                                   ByVal SourceFileName As String, _
                                   ByVal CaseStatus As String) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim FolderName As String
Dim SpFileName As String
Dim FinalFileName As String
Dim Extension As String
Dim DropboxPath As String
Dim ScanFolder As String
Dim ScanDropboxPath As String

Dim tempDir As String
Dim pickedPath As String
Dim fDialog As Office.FileDialog
Dim varFile As Variant
Dim pickedExists As Boolean

Dim sourceSize As Currency
Dim fso As Object
Dim uploadOk As Boolean
Dim spOk As Boolean
Dim deleteOk As Boolean

    rv = False

    ' --- 1. Resolve Dropbox destination folder + filename -----------------
    If CaseStatus = "Closed" Then
        FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
    Else
        FolderName = GetDocumentFolderName(CaseID, DocumentType)
    End If

    If pcaempty(FolderName) Then
        MsgBox "Could not resolve the Dropbox destination folder for this case.", _
               vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    SpFileName = GetDocumentFileName(CaseID, DocumentType)
    If InStrRev(SourceFileName, ".") > 0 Then
        Extension = Right$(SourceFileName, Len(SourceFileName) - InStrRev(SourceFileName, "."))
        SpFileName = SpFileName & "." & Extension
    End If

    ' --- 2. Override-with-confirmation Save-As ----------------------------
    ' Office.FileDialog can't show Dropbox paths, so we root it in
    ' %TEMP%\TBCMS\ and use only the filename portion the user submits.
    tempDir = Environ$("TEMP") & "\TBCMS"
    On Error Resume Next
    MkDir tempDir    ' no-op if it exists; CleanupTempFiles sweeps it
    Err.Clear
    On Error GoTo Err_Handler

    Set fDialog = Application.FileDialog(msoFileDialogSaveAs)
    pickedExists = False
    With fDialog
        .AllowMultiSelect = False
        .Title = "Save to Dropbox case folder"
        .InitialFileName = tempDir & "\" & SpFileName
        If .show = True Then
            For Each varFile In .SelectedItems
                pickedPath = varFile
                pickedExists = True
            Next
        End If
    End With

    If Not pickedExists Then
        ' User cancelled the dialog — not an error, just abort
        rv = True
        GoTo Exit_Handler
    End If

    FinalFileName = Mid$(pickedPath, InStrRev(pickedPath, "\") + 1)
    If FinalFileName <> SpFileName Then
        If MsgBox("You have changed the suggested filename from:" & vbCrLf & vbCrLf & _
                  "    " & SpFileName & vbCrLf & vbCrLf & _
                  "to:" & vbCrLf & vbCrLf & _
                  "    " & FinalFileName & vbCrLf & vbCrLf & _
                  "Save with the new name?", _
                  vbYesNo + vbQuestion, "TB CMS — confirm filename change") = vbNo Then
            rv = True
            GoTo Exit_Handler
        End If
    End If

    ' --- 3. Compute Dropbox destination path + route by size --------------
    DropboxPath = FolderName
    If Right$(DropboxPath, 1) <> "/" Then DropboxPath = DropboxPath & "/"
    DropboxPath = DropboxPath & FinalFileName

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(SourceFileName) Then
        MsgBox "Source file not found: " & SourceFileName, vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If
    sourceSize = CCur(fso.GetFile(SourceFileName).Size)

    If sourceSize > 157286400 Then    ' 150 MB
        uploadOk = DropboxService.UploadLargeFile(SourceFileName, DropboxPath, CaseID, DocumentType)
    Else
        uploadOk = DropboxService.UploadFile(SourceFileName, DropboxPath, CaseID, DocumentType)
    End If

    If Not uploadOk Then
        ' Upload itself failed (e.g., gated by ALLOW_DROPBOX_WRITES). No
        ' orphan to compensate — the upload never landed.
        MsgBox "Upload to Dropbox failed. Check tblDropboxAuditLog for details.", _
               vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    ' --- 4. Register via SaveCaseDocument (with G13 token guard) ----------
    spOk = SaveCaseDocument(CaseID, DocumentType, DropboxPath)

    If Not spOk Then
        ' Orphan-file compensation: upload succeeded but SQL register failed.
        ' Try to delete the upload. If delete also fails, queue the orphan.
        deleteOk = DropboxService.DeleteFile(DropboxPath, CaseID, DocumentType)
        If Not deleteOk Then
            QueueOrphanFile DropboxPath, "ScanSave", CaseID, DocumentType, _
                "SaveCaseDocument returned False (see tblDropboxLog)", _
                "DropboxService.DeleteFile returned False"
        End If
        MsgBox "The document was uploaded to Dropbox but could not be " & _
               "registered in TBCMS. " & _
               IIf(deleteOk, "The upload has been removed.", _
                             "The upload could NOT be removed and has been " & _
                             "queued for IT-admin cleanup (tblDropboxOrphanQueue).") & _
               vbCrLf & vbCrLf & "See tblDropboxLog / tblDropboxAuditLog for details.", _
               vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    ' --- 5. Closed Final special case: optional copy to ClosedFileScan ----
    If DocumentType = "Closed Final" Then
        If MsgBox("Do you want to save the file in the Closed File Scans directory?", _
                  vbYesNo + vbQuestion, "TB CMS") = vbYes Then
            ScanFolder = GetClosedFileScanFolderName(CaseID, "General")
            If LenB(ScanFolder) > 0 Then
                If Right$(ScanFolder, 1) <> "/" Then ScanFolder = ScanFolder & "/"
                ScanDropboxPath = ScanFolder & FinalFileName
                ' Copy from the just-uploaded source. Failure here is logged
                ' to tblDropboxAuditLog by DropboxService.CopyFile but does
                ' NOT roll back the primary upload (the Closed File Scan
                ' copy is a convenience, not the source of truth).
                DropboxService.CopyFile DropboxPath, ScanDropboxPath, CaseID, DocumentType
            End If
        End If
    End If

    rv = True
Exit_Handler:
    SaveScannedFileAs = rv
    Exit Function
Err_Handler:
    rv = False
    If Err.Number = vbObjectError + 6001 Then
        ' GuardWritesEnabled — test environment, writes disabled.
        MsgBox "Dropbox writes are disabled in this test build " & _
               "(ALLOW_DROPBOX_WRITES = False). No file was uploaded.", _
               vbInformation, "TB CMS — test environment"
    Else
        foo = pcaStdErrMsg(Err, Error)
    End If
    Resume Exit_Handler
End Function


' ----------------------------------------------------------------------------
' Inserts a row into tblDropboxOrphanQueue when an upload-then-SP-then-
' compensating-delete sequence fails at the delete step. IT admin drains
' the queue per the runbook (see Phase 2 — tblDropboxOrphanQueue schema +
' Phase 5 — orphan-file compensation policy).
' Failure to queue the orphan is logged to tblDropboxLog but does not
' propagate further — the caller has already shown the user the upload-
' succeeded-register-failed-delete-failed error and continuing to surface
' "couldn't even queue the orphan" doesn't help.
Private Sub QueueOrphanFile(ByVal dropboxPath As String, _
                             ByVal workflowName As String, _
                             ByVal caseID As Variant, _
                             ByVal documentType As String, _
                             ByVal originalSPError As String, _
                             ByVal compensatingDeleteError As String)
    On Error GoTo HandleError
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    Dim sql As String
    Dim caseSql As String
    If IsMissing(caseID) Or IsNull(caseID) Then
        caseSql = "NULL"
    Else
        caseSql = CStr(CLng(caseID))
    End If

    sql = "INSERT INTO dbo.tblDropboxOrphanQueue " & _
          "(DropboxAccountEmail, OrphanDropboxPath, WorkflowName, CaseID, " & _
          " DocumentType, OriginalSPError, CompensatingDeleteError) VALUES (" & _
          pcaAddQuotes(DropboxService.GetDropboxAccountEmail()) & ", " & _
          pcaAddQuotes(dropboxPath) & ", " & _
          pcaAddQuotes(workflowName) & ", " & _
          caseSql & ", " & _
          pcaAddQuotes(documentType) & ", " & _
          pcaAddQuotes(originalSPError) & ", " & _
          pcaAddQuotes(compensatingDeleteError) & ")"

    cn.Execute sql
    cn.Close
    Exit Sub
HandleError:
    On Error Resume Next
    DropboxService.LogLocal "QueueOrphanFile", "Error", _
        "Failed to insert orphan-queue row for " & dropboxPath & ": Err=" & _
        Err.Number & " " & Err.Description
    Err.Clear
End Sub


Public Function SaveCaseDocument(ByVal CaseID As Integer, ByVal DocumentType As String, ByVal DocumentFileName As String) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim cn As ADODB.Connection
Dim sql As String

    rv = False
    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spSaveCaseDocument "
    sql = sql & "@CaseID = " & CaseID
    sql = sql & ",@DocumentType = " & pcaAddQuotes(DocumentType)
    sql = sql & ",@DocumentName = " & pcaAddQuotes(DocumentFileName)

    cn.Execute sql
    rv = True
Exit_Handler:
    SaveCaseDocument = rv
    Exit Function
Err_Handler:
    rv = False
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


Public Function GetCaseDocument(ByVal CaseID As Integer, ByVal DocumentType) As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    rv = ""
    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetCaseDocument "
    sql = sql & "@CaseID = " & CaseID
    sql = sql & ",@DocumentType = " & pcaAddQuotes(DocumentType)

    Set rs = cn.Execute(sql)
    If Not rs.EOF Then
        rv = rs("DocumentFileName")
    End If
Exit_Handler:
    GetCaseDocument = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


' ============================================================================
' LEGACY (pre-Phase 4a) — OpenDocumentFolder
' ----------------------------------------------------------------------------
' Preserved below as a commented block so we can roll back to S:\ behavior
' if 4a needs to be reverted. To roll back:
'   1. Comment out the active OpenDocumentFolder below (prepend ' to every
'      line of the function body — VBE: select lines, Edit > Comment Block).
'   2. Uncomment the LEGACY version (remove the leading "' " from each line).
'   3. Re-import this file via VBE > File > Import File (replace).
' ----------------------------------------------------------------------------
' Public Function OpenDocumentFolder(ByVal CaseID As Variant, ByVal DocumentType) As Boolean
' On Error GoTo Err_Handler
' Dim rv As Boolean
' Dim FolderName As String
'
'     rv = False
'
'     If pcaempty(CaseID) Then
'         MsgBox "Please select a case before proceeding...", , "TB CMS"
'     Else
'         If GetCaseClosedStatus(CaseID) Then
'             FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
'         Else
'             FolderName = GetDocumentFolderName(CaseID, DocumentType)
'         End If
'
'         If Not FolderExistsCreate(FolderName, False) Then
'             If MsgBox(FolderName & " Folder for this case doesn't exists.  Do you want to create it?", vbYesNo, "TB CMS") = vbYes Then
'                 If FolderExistsCreate(FolderName, True) Then
'                     MsgBox "Document folder is created", , "TB CMS"
'                     Call OpenFileDialog("Case Document", FolderName, "")
'                 End If
'             End If
'         Else
'             Call OpenFileDialog("Case Document", FolderName, "")
'         End If
'     End If
'     rv = True
' Exit_Handler:
'     OpenDocumentFolder = rv
'     Exit Function
' Err_Handler:
'     rv = False
'     Resume Exit_Handler
' End Function
' ============================================================================


' ============================================================================
' LEGACY (pre-Phase 4b — Phase 4a version) — OpenDocumentFolder
' ----------------------------------------------------------------------------
' Preserved for rollback. This is the Phase 4a body (browser-only, no
' Explorer routing). The Phase 4b active version below tries Explorer first
' and falls back to the browser. To roll back: comment-out the active body,
' uncomment this LEGACY block, re-import.
' ----------------------------------------------------------------------------
' Public Function OpenDocumentFolder(ByVal CaseID As Variant, ByVal DocumentType As Variant) As Boolean
' On Error GoTo Err_Handler
' Dim rv As Boolean
' Dim FolderName As String
' Dim webUrl As String
' Dim pathForCheck As String
' Dim found As Boolean
' Dim errDetail As String
' Dim mdJson As String
' Dim apiOk As Boolean
'     rv = False
'     If pcaempty(CaseID) Then
'         MsgBox "Please select a case before proceeding...", , "TB CMS"
'     Else
'         If GetCaseClosedStatus(CaseID) Then
'             FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
'         Else
'             FolderName = GetDocumentFolderName(CaseID, DocumentType)
'         End If
'         If pcaempty(FolderName) Then
'             MsgBox "Could not resolve the document folder for this case.", vbExclamation, "TB CMS"
'         ElseIf Left$(FolderName, 1) <> "/" Then
'             MsgBox "Folder path is not a Dropbox path (got: " & FolderName & "). Contact IT...", vbExclamation, "TB CMS"
'         Else
'             pathForCheck = FolderName
'             If Right$(pathForCheck, 1) = "/" Then
'                 pathForCheck = Left$(pathForCheck, Len(pathForCheck) - 1)
'             End If
'             apiOk = DropboxService.GetMetadata(pathForCheck, found, errDetail, mdJson)
'             If apiOk And Not found Then
'                 MsgBox "The Dropbox folder for this case doesn't exist yet...", vbExclamation, "TB CMS"
'             Else
'                 webUrl = "https://www.dropbox.com/work" & FolderName
'                 Application.FollowHyperlink webUrl
'             End If
'         End If
'     End If
'     rv = True
' Exit_Handler:
'     OpenDocumentFolder = rv
'     Exit Function
' Err_Handler:
'     rv = False
'     foo = pcaStdErrMsg(Err, Error)
'     Resume Exit_Handler
' End Function
' ============================================================================


' --- Phase 4b rewire --------------------------------------------------------
' Folder open: route through Windows Explorer against the local
' Dropbox-synced folder when the desktop client has resolved the team root
' on disk. Falls back to the Dropbox web URL when:
'   - The desktop client isn't installed/signed in (m_LocalSyncedRoot empty)
'   - The local synced folder doesn't exist on disk (case never synced)
'
' Explorer routing avoids the "Unsupported path provided" warning banner
' Dropbox shows for any /home/ or /work/ deep-link into team-namespace
' content, and gives attorneys the familiar File Explorer UX they used in
' the S:\ days.
Public Function OpenDocumentFolder(ByVal CaseID As Variant, ByVal DocumentType As Variant) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim FolderName As String
Dim webUrl As String
Dim localPath As String
Dim pathForCheck As String
Dim found As Boolean
Dim errDetail As String
Dim mdJson As String
Dim apiOk As Boolean

    rv = False

    If pcaempty(CaseID) Then
        MsgBox "Please select a case before proceeding...", , "TB CMS"
    Else
        If GetCaseClosedStatus(CaseID) Then
            FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
        Else
            FolderName = GetDocumentFolderName(CaseID, DocumentType)
        End If

        If pcaempty(FolderName) Then
            MsgBox "Could not resolve the document folder for this case.", vbExclamation, "TB CMS"
        ElseIf Left$(FolderName, 1) <> "/" Then
            MsgBox "Folder path is not a Dropbox path (got: " & FolderName & "). " & _
                   "Contact IT — this indicates a configuration issue.", vbExclamation, "TB CMS"
        Else
            ' Try the local-synced folder first (best UX — no browser banner).
            localPath = DropboxService.DropboxPathToLocalPath(FolderName)
            If LenB(localPath) > 0 Then
                ' Strip trailing slash that DropboxPathToLocalPath may produce
                ' from the trailing '/' on FolderName.
                If Right$(localPath, 1) = "\" Then
                    localPath = Left$(localPath, Len(localPath) - 1)
                End If
                If Dir$(localPath, vbDirectory) <> "" Then
                    ' Local folder is synced and exists — open Explorer there.
                    Shell "explorer.exe """ & localPath & """", vbNormalFocus
                    rv = True
                    GoTo Exit_Handler
                End If
            End If

            ' Fallback path: pre-check existence in Dropbox, then web URL.
            pathForCheck = FolderName
            If Right$(pathForCheck, 1) = "/" Then
                pathForCheck = Left$(pathForCheck, Len(pathForCheck) - 1)
            End If

            apiOk = DropboxService.GetMetadata(pathForCheck, found, errDetail, mdJson)

            If apiOk And Not found Then
                MsgBox "The Dropbox folder for this case doesn't exist yet:" & vbCrLf & vbCrLf & _
                       FolderName & vbCrLf & vbCrLf & _
                       "This typically means no documents have been saved for this " & _
                       "case + document-type combination yet. The folder will be " & _
                       "created automatically when the first document is saved.", _
                       vbExclamation, "TB CMS"
            Else
                ' Either GetMetadata confirmed found, or there was a transport
                ' failure — open the Dropbox web URL. /work/ is the team-content
                ' prefix; expect a cosmetic "Unsupported path provided" banner
                ' from Dropbox for team-namespace deep links.
                webUrl = "https://www.dropbox.com/work" & FolderName
                Application.FollowHyperlink webUrl
            End If
        End If
    End If
    rv = True
Exit_Handler:
    OpenDocumentFolder = rv
    Exit Function
Err_Handler:
    rv = False
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


' ============================================================================
' LEGACY (pre-Phase 4a) — OpenDocumentFile
' ----------------------------------------------------------------------------
' Preserved below as a commented block so we can roll back to S:\ behavior
' if 4a needs to be reverted. To roll back: see the swap procedure documented
' above the LEGACY OpenDocumentFolder block.
' ----------------------------------------------------------------------------
' Public Function OpenDocumentFile(ByVal CaseID As Variant, ByVal DocumentType As String) As Boolean
' On Error GoTo Err_Handler
' Dim rv As Boolean
' Dim DocumentFileName As String
'     rv = False
'
'     If pcaempty(CaseID) Then
'         MsgBox "Please select a case before proceeding...", , "TB CMS"
'     Else
'         DocumentFileName = GetCaseDocument(CaseID, DocumentType)
'         'check if the document exists
'         If Not pcaempty(DocumentFileName) And Dir(DocumentFileName) <> "" Then
'             Application.FollowHyperlink DocumentFileName
'         Else
'             MsgBox DocumentType & " is not found", vbExclamation, "TB CMS"
'         End If
'     End If
'     rv = True
' Exit_Handler:
'     OpenDocumentFile = rv
'     Exit Function
' Err_Handler:
'     rv = False
'     Resume Exit_Handler
' End Function
' ============================================================================


' --- Phase 4a rewire --------------------------------------------------------
' Document open: route through DropboxService.OpenDocument, which downloads
' the file to %TEMP%\TBCMS\<GUID>_<filename> and hands the local path to
' Application.FollowHyperlink so the document opens in its native app.
'
' Behavioral change vs. legacy S:\ flow (call out in user runbook): edits
' made in the launched app are NOT auto-re-uploaded. Users must save
' changes back through the Save flow. The legacy UNC-share flow had
' implicit save-on-close semantics; this does not.
Public Function OpenDocumentFile(ByVal CaseID As Variant, ByVal DocumentType As String) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim DocumentFileName As String
Dim localPath As String

    rv = False

    If pcaempty(CaseID) Then
        MsgBox "Please select a case before proceeding...", , "TB CMS"
    Else
        DocumentFileName = GetCaseDocument(CaseID, DocumentType)
        If pcaempty(DocumentFileName) Then
            MsgBox DocumentType & " is not on file for this case.", vbExclamation, "TB CMS"
        ElseIf Left$(DocumentFileName, 1) <> "/" Then
            MsgBox "Document path is not a Dropbox path (got: " & DocumentFileName & "). " & _
                   "Contact IT — this indicates a configuration issue.", vbExclamation, "TB CMS"
        Else
            localPath = DropboxService.OpenDocument(DocumentFileName)
            If pcaempty(localPath) Then
                MsgBox DocumentType & " could not be opened from Dropbox. " & _
                       "Check tblDropboxLog for details, or contact IT.", _
                       vbExclamation, "TB CMS"
            End If
        End If
    End If
    rv = True
Exit_Handler:
    OpenDocumentFile = rv
    Exit Function
Err_Handler:
    rv = False
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


' ============================================================================
' LEGACY (pre-Phase 4d) — MoveDocumentByCaseStatus
' ----------------------------------------------------------------------------
' Preserved for rollback. Used FSO.CopyFolder + FSO.DeleteFolder + the
' legacy spMoveDocumentFolder(@CaseID, @CaseStatus) signature. The legacy
' SP is gone after Phase 4c — re-running this body against the post-4c
' database would fail at the spMoveDocumentFolder call ("no parameter
' @CaseStatus"). Rolling back to this body also requires reverting the
' Phase 4c SQL.
' ============================================================================
' Public Function MoveDocumentByCaseStatus(ByVal CaseID As Variant, ByVal CaseStatus As String) As Boolean
' On Error GoTo Err_Handler
' Dim rv As Boolean
' Dim SourceFolder As String
' Dim TargetFolder As String
' Dim FSO As Object
' Dim cn As ADODB.Connection
' Dim sql As String
' Dim i As Integer
' Dim LArray() As String
'     rv = False
'     If CaseStatus = "Closed" Then
'         SourceFolder = GetDocumentFolderName(CaseID, "General")
'         TargetFolder = GetClosedDocumentFolderName(CaseID, "General")
'     Else
'         SourceFolder = GetClosedDocumentFolderName(CaseID, "Init Intake, Notes, Documents")
'         TargetFolder = GetDocumentFolderName(CaseID, "Init Intake, Notes, Documents")
'     End If
'     LArray = Split(TargetFolder, "\")
'     i = 0
'     Do While (LArray(i) <> "")
'             i = i + 1
'     Loop
'     TargetFolder = Left(TargetFolder, Len(TargetFolder) - Len(LArray(i - 1)) - 1)
'     If Right(SourceFolder, 1) = "\" Then
'         SourceFolder = Left(SourceFolder, Len(SourceFolder) - 1)
'     End If
'     If Right(TargetFolder, 1) = "\" Then
'         TargetFolder = Left(TargetFolder, Len(TargetFolder) - 1)
'     End If
'     Set FSO = CreateObject("scripting.filesystemobject")
'     If Not FSO.FolderExists(SourceFolder) Then
'         MsgBox "Source folder doesn't exists..."
'     Else
'         If Not FolderExistsCreate(TargetFolder, True) Then
'             MsgBox "Failed to create target folder"
'         Else
'             FSO.CopyFolder Source:=SourceFolder, Destination:=TargetFolder
'             FSO.DeleteFolder SourceFolder
'             Set cn = New ADODB.Connection
'             cn.Open PcaGetConnnectionString
'             sql = "exec spMoveDocumentFolder @CaseID=" & CaseID & ",@CaseStatus=" & pcaAddQuotes(CaseStatus)
'             cn.Execute sql
'             rv = True
'         End If
'     End If
' Exit_Handler:
'     MoveDocumentByCaseStatus = rv
'     Exit Function
' Err_Handler:
'     ...
' End Function
' ============================================================================


' --- Phase 4d rewire --------------------------------------------------------
' Case close/reopen folder move (Phase 5 step 4 in the plan). The two
' callers in frmClientLedger — cmdCloseCase_Click and cmdReopenCase_Click —
' keep their current signatures and pass "Closed" or "Open".
'
' Sequence:
'   1. Resolve the source + destination Dropbox folder paths via the
'      post-4c path-building SPs. For Closed: source = open-case folder,
'      destination = closed-case folder. For Open: reversed.
'   2. Call DropboxService.MoveFile(@OldFolderPath, @NewFolderPath). Works
'      at folder granularity (Dropbox move_v2 handles files OR folders).
'   3. On move success: call the new G2 spMoveDocumentFolder with explicit
'      @OldFolderPath / @NewFolderPath. Three branches:
'        - SP returns both rowcounts > 0       → commit (normal case)
'        - SP returns exactly one rowcount = 0 → warn + commit (case has
'                                                docs but no scans, or
'                                                vice versa)
'        - SP throws (both-zero or runtime err)→ compensate: reverse the
'                                                Dropbox move, log Failure
'   4. On Dropbox move failure (conflict, transport error): no SP call,
'      surface error.
'
' Test-environment behavior: DropboxService.MoveFile hits GuardWritesEnabled
' and raises immediately. The error propagates here; we surface a generic
' "move failed" message and return False.
Public Function MoveDocumentByCaseStatus(ByVal CaseID As Variant, _
                                          ByVal CaseStatus As String) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim OldFolderPath As String
Dim NewFolderPath As String
Dim moveOk As Boolean
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim caseDocsUpdated As Long
Dim scansUpdated As Long
Dim spErrNum As Long
Dim spErrDesc As String
Dim revertOk As Boolean

    rv = False

    ' --- 1. Resolve old + new folder paths --------------------------------
    If CaseStatus = "Closed" Then
        OldFolderPath = GetDocumentFolderName(CaseID, "General")
        NewFolderPath = GetClosedDocumentFolderName(CaseID, "General")
    Else
        OldFolderPath = GetClosedDocumentFolderName(CaseID, "Init Intake, Notes, Documents")
        NewFolderPath = GetDocumentFolderName(CaseID, "Init Intake, Notes, Documents")
    End If

    If pcaempty(OldFolderPath) Or pcaempty(NewFolderPath) Then
        MsgBox "Could not resolve the source or target Dropbox folder paths " & _
               "for case " & CaseID & ".", vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    ' Strip trailing '/' for DropboxService.MoveFile — Dropbox move_v2
    ' expects no trailing slash on folder paths.
    If Right$(OldFolderPath, 1) = "/" Then _
        OldFolderPath = Left$(OldFolderPath, Len(OldFolderPath) - 1)
    If Right$(NewFolderPath, 1) = "/" Then _
        NewFolderPath = Left$(NewFolderPath, Len(NewFolderPath) - 1)

    ' --- 2. Dropbox folder move -------------------------------------------
    moveOk = DropboxService.MoveFile(OldFolderPath, NewFolderPath, CaseID, "General")
    If Not moveOk Then
        ' Move itself failed (e.g., gated, conflict, transport error). No
        ' SP call. tblDropboxAuditLog has the detail.
        MsgBox "Dropbox folder move failed for case " & CaseID & ". " & _
               "Check tblDropboxAuditLog for the underlying error.", _
               vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    ' --- 3. Call new G2 spMoveDocumentFolder ------------------------------
    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = "EXEC dbo.spMoveDocumentFolder " & _
          "@CaseID = " & CaseID & ", " & _
          "@OldFolderPath = " & pcaAddQuotes(OldFolderPath) & ", " & _
          "@NewFolderPath = " & pcaAddQuotes(NewFolderPath)

    On Error Resume Next
    Set rs = cn.Execute(sql)
    spErrNum = Err.Number
    spErrDesc = Err.Description
    On Error GoTo Err_Handler
    Err.Clear

    If spErrNum <> 0 Then
        ' SP threw — most likely the G2 both-zero-rowcount throw (51000),
        ' meaning @OldFolderPath doesn't prefix-match any tblCaseDocuments
        ' or tblScans rows for this case. Compensate by reversing the
        ' Dropbox move.
        revertOk = DropboxService.MoveFile(NewFolderPath, OldFolderPath, CaseID, "General")
        cn.Close
        If revertOk Then
            MsgBox "Case " & CaseID & " could not be " & LCase$(CaseStatus) & _
                   ". The SQL ledger has no record under the source folder, " & _
                   "so the Dropbox move has been reverted." & vbCrLf & vbCrLf & _
                   "Underlying error: " & spErrDesc & vbCrLf & vbCrLf & _
                   "Contact IT.", vbExclamation, "TB CMS"
        Else
            ' Critical: forward move succeeded but reverse move failed too.
            ' Case is in an inconsistent state. Per plan G2: log to audit,
            ' surface stronger user message, do NOT auto-retry.
            DropboxService.LogAuditEvent "Move", "Failure", CaseID, "General", _
                NewFolderPath & " -> " & OldFolderPath, _
                "CRITICAL: SP error '" & spErrDesc & "' + reverse-move failed too"
            MsgBox "Case " & CaseID & " is in an INCONSISTENT state: the " & _
                   "Dropbox folder was moved to:" & vbCrLf & "    " & NewFolderPath & vbCrLf & _
                   "but the SQL ledger has no record there AND the reverse " & _
                   "move also failed." & vbCrLf & vbCrLf & _
                   "DO NOT retry. Contact IT IMMEDIATELY.", _
                   vbCritical, "TB CMS — CRITICAL"
        End If
        rv = True
        GoTo Exit_Handler
    End If

    ' SP returned a recordset with CaseDocumentsUpdated + ScansUpdated.
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            caseDocsUpdated = CLng(pcaConvertNulls(rs("CaseDocumentsUpdated"), 0))
            scansUpdated = CLng(pcaConvertNulls(rs("ScansUpdated"), 0))
        End If
        rs.Close
        Set rs = Nothing
    End If
    cn.Close
    Set cn = Nothing

    ' Exactly-one-zero rowcount is a warn-and-accept case (common when a
    ' case has tblCaseDocuments rows but no tblScans rows, or vice versa).
    If (caseDocsUpdated = 0) Xor (scansUpdated = 0) Then
        DropboxService.LogAuditEvent "Move", "Success", CaseID, "General", _
            OldFolderPath & " -> " & NewFolderPath, _
            "one-table-zero-rowcount: " & _
            IIf(caseDocsUpdated = 0, "tblCaseDocuments", "tblScans") & " had no rows"
    End If

    rv = True
Exit_Handler:
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
        On Error GoTo 0
    End If
    If Not cn Is Nothing Then
        On Error Resume Next
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
        On Error GoTo 0
    End If
    MoveDocumentByCaseStatus = rv
    Exit Function
Err_Handler:
    rv = False
    If Err.Number = vbObjectError + 6001 Then
        ' GuardWritesEnabled — test environment, writes disabled.
        MsgBox "Dropbox writes are disabled in this test build " & _
               "(ALLOW_DROPBOX_WRITES = False). The case was not moved.", _
               vbInformation, "TB CMS — test environment"
    Else
        foo = pcaStdErrMsg(Err, Error)
    End If
    Resume Exit_Handler
End Function


' ============================================================================
' LEGACY (pre-Phase 4d) — CopyDocumentToClosedFileScan
' ============================================================================
' Public Function CopyDocumentToClosedFileScan(ByVal CaseID As Variant) As Boolean
' On Error GoTo Err_Handler
' Dim rv As Boolean
' Dim SourceFolder As String
' Dim TargetFolder As String
' Dim FSO As Object
' Dim i As Integer
' Dim LArray() As String
'     rv = False
'     SourceFolder = GetDocumentFolderName(CaseID, "General")
'     TargetFolder = GetClosedFileScanFolderName(CaseID, "General")
'     LArray = Split(TargetFolder, "\")
'     i = 0
'     Do While (LArray(i) <> "")
'             i = i + 1
'     Loop
'     TargetFolder = Left(TargetFolder, Len(TargetFolder) - Len(LArray(i - 1)) - 1)
'     If Right(SourceFolder, 1) = "\" Then
'         SourceFolder = Left(SourceFolder, Len(SourceFolder) - 1)
'     End If
'     If Right(TargetFolder, 1) = "\" Then
'         TargetFolder = Left(TargetFolder, Len(TargetFolder) - 1)
'     End If
'     Set FSO = CreateObject("scripting.filesystemobject")
'     If Not FSO.FolderExists(SourceFolder) Then
'         MsgBox "Source folder doesn't exists..."
'     Else
'         If Not FolderExistsCreate(TargetFolder, True) Then
'             MsgBox "Fail to create target folder"
'         Else
'             FSO.CopyFolder Source:=SourceFolder, Destination:=TargetFolder
'             rv = True
'         End If
'     End If
' Exit_Handler:
'     CopyDocumentToClosedFileScan = rv
'     Exit Function
' Err_Handler:
'     foo = pcaStdErrMsg(Err, Error)
'     rv = False
'     Resume Exit_Handler
' End Function
' ============================================================================


' --- Phase 4d rewire --------------------------------------------------------
' Optional first step of case-close (Phase 5 step 4): copy the open-case
' folder into the Closed File Scans area before the main move. Caller in
' frmClientLedger.cmdCloseCase_Click invokes this when the user confirms
' the "save in Closed File Scans?" prompt.
'
' Path construction:
'   Source: GetDocumentFolderName(CaseID, "General") — open-case folder
'   Target parent: GetClosedFileScanFolderName(CaseID, "General") — the
'                  TB/<Yr>/ folder under /Company/Closed File Scans
'   Destination: <target parent> + basename(source) — places a copy of
'                the case folder under TB/<Yr>/
'
' Dropbox copy_v2 with autorename=false; conflict surfaces an error.
Public Function CopyDocumentToClosedFileScan(ByVal CaseID As Variant) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim SourceFolder As String
Dim TargetParent As String
Dim DestinationFolder As String
Dim sourceBaseName As String

    rv = False

    SourceFolder = GetDocumentFolderName(CaseID, "General")
    TargetParent = GetClosedFileScanFolderName(CaseID, "General")

    If pcaempty(SourceFolder) Or pcaempty(TargetParent) Then
        MsgBox "Could not resolve source or target Dropbox folder for the " & _
               "Closed File Scans copy.", vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    ' Strip trailing '/' to make basename + concatenation predictable.
    If Right$(SourceFolder, 1) = "/" Then _
        SourceFolder = Left$(SourceFolder, Len(SourceFolder) - 1)
    If Right$(TargetParent, 1) = "/" Then _
        TargetParent = Left$(TargetParent, Len(TargetParent) - 1)

    ' Basename: segment after the last '/'.
    sourceBaseName = Mid$(SourceFolder, InStrRev(SourceFolder, "/") + 1)
    If LenB(sourceBaseName) = 0 Then
        MsgBox "Could not derive a folder name from " & SourceFolder & ".", _
               vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    DestinationFolder = TargetParent & "/" & sourceBaseName

    ' Dropbox copy_v2 — autorename=false. Failure (gated kill-switch, path
    ' conflict, transport error) is logged to tblDropboxAuditLog by
    ' DropboxService.CopyFile.
    If Not DropboxService.CopyFile(SourceFolder, DestinationFolder, CaseID, "General") Then
        MsgBox "Failed to copy the case folder to Closed File Scans. " & _
               "Check tblDropboxAuditLog for the underlying error.", _
               vbExclamation, "TB CMS"
        rv = True
        GoTo Exit_Handler
    End If

    rv = True
Exit_Handler:
    CopyDocumentToClosedFileScan = rv
    Exit Function
Err_Handler:
    rv = False
    If Err.Number = vbObjectError + 6001 Then
        ' GuardWritesEnabled — test environment, writes disabled.
        MsgBox "Dropbox writes are disabled in this test build " & _
               "(ALLOW_DROPBOX_WRITES = False). No copy was made to " & _
               "Closed File Scans.", vbInformation, "TB CMS — test environment"
    Else
        foo = pcaStdErrMsg(Err, Error)
    End If
    Resume Exit_Handler
End Function


Public Function GetCaseClosedStatus(ByVal CaseID As Integer) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    rv = False
    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetCaseClosedStatus "
    sql = sql & "@CaseID = " & CaseID

    Set rs = cn.Execute(sql)

    If Not rs.EOF Then
        rv = rs("Closed")
    Else
        rv = False
    End If
Exit_Handler:
    GetCaseClosedStatus = rv
    Exit Function
Err_Handler:
    rv = False
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function


Public Function GetIntakeDocumentFileName(ByVal IntakeID As Long) As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "exec spGetIntakeDocumentFileName "
    sql = sql & "@IntakeID = " & IntakeID

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = pcaConvertNulls(rs("FileName"), "")
    Else
        rv = ""
    End If
Exit_Handler:
    GetIntakeDocumentFileName = rv
    Exit Function
Err_Handler:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume Exit_Handler
End Function
