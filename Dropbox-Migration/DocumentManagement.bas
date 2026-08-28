Attribute VB_Name = "DocumentManagement"
Option Compare Database
Option Explicit
Dim foo

' =============================================================================
' Phase 4 rewire status (last updated 2026-06-03):
'
'   Phase 4a — Read-flow rewire (DONE):
'     OpenDocumentFile     -> DropboxService.OpenDocument
'                             (download to %TEMP%\TBCMS\ + native-app launch)
'     OpenDocumentFolder   -> Explorer-first open (Phase 4e, G24): opens the
'                             Dropbox-client local mount in Windows Explorer
'                             when available, else falls back to the Dropbox
'                             web URL via explorer.exe (G27). Also restores
'                             legacy create-on-demand — on a confirmed-missing
'                             folder it prompts and calls
'                             DropboxService.CreateFolder (write op, gated by
'                             ALLOW_DROPBOX_WRITES) before opening
'
'   Phase 4b — Config layer + local-synced-root routing (PARTLY REVIVED in 4e):
'     The Dropbox desktop client was originally not a hard deployment
'     prerequisite; the code degrades gracefully without it. As of Phase 4e the
'     local-synced-root helpers are wired back into folder-open (Explorer-first);
'     the remaining helpers stay dormant.
'     2026-08-28: firm policy now assumes the desktop client is installed on
'     every workstation, reversing the "not a hard prerequisite" premise below.
'     Comments in this module still describe the graceful-degradation design;
'     that fallback behavior is retained defensively but should no longer be
'     the expected runtime path. Current behavior:
'       OpenDocumentFolder   -> Explorer (local mount) when the desktop client
'                               is present and the folder has synced, else the
'                               Dropbox web URL. Web deep-links to /work/ team-
'                               namespace paths still show Dropbox's cosmetic
'                               "Unsupported path provided" banner.
'                               (2026-06-03: create-on-demand re-added on top
'                               of the open — see Phase 4a note above.)
'       GetDocumentRootFolder -> reads tblDropboxRootConfig.TeamRootPath +
'                                DropboxService.DropboxPathToLocalPath.
'                                Zero callers (kept for contract stability);
'                                returns "" without the desktop client.
'       GetScannerFolder     -> returns "" by design. Office.FileDialog
'                               opens at the user's last-used folder.
'                               Two callers: Intakes.cmdScan_Click +
'                               frmClientLedger scan flow.
'     DropboxService.DropboxPathToLocalPath + ResolveLocalSyncedRoot are now
'     invoked again by OpenDocumentFolder (Phase 4e); LocalPathToDropboxPath
'     remains in place for future ingest use.
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
' 2026-08-28: under current firm policy the desktop client is expected on
' every workstation, so this "" case should be rare in practice — this
' function has zero live callers regardless (see below).
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


' --- Phase 4 follow-up: returns "" by design -------------------------------
' Office.FileDialog cannot navigate Dropbox API paths, and at the time this
' was written the firm had opted not to require the Dropbox desktop client —
' so there was no Windows path to hand back. Callers (Intakes.cmdScan_Click +
' frmClientLedger scan flow) pass "" to FileDialog.InitialFileName, which
' opens at the user's last-used folder.
'
' 2026-08-28: firm policy now assumes the desktop client is installed on
' every workstation, which reopens the option this comment describes below
' (a real Windows path via DropboxService.DropboxPathToLocalPath /
' ResolveLocalSyncedRoot). Left as "" for now — not changed as part of this
' pass — but this is the function to revisit if scanner-folder routing is
' picked up as follow-up work.
'
' Kept as a function (rather than deleted) so callers continue to compile.
' If a local-path strategy is reintroduced later (e.g. a Windows path
' column in tblDropboxRootConfig), wire it back through this function.
Public Function GetScannerFolder() As String
    GetScannerFolder = ""
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


' --- Phase 4 follow-up: folder open + create-on-demand ----------------------
' STALE COMMENT, corrected 2026-08-28: the paragraph below ("do NOT route
' through Windows Explorer...") described the Phase 4b/4d web-only design.
' It was left unupdated when Phase 4e (2026-06-29) rewired this function to
' call OpenFolderInExplorerOrWeb (line ~1050 below), which DOES try Explorer
' first via the local-synced root and falls back to the web URL only if the
' desktop client/local mount isn't available. That fallback is exactly what
' the "do NOT route through Explorer" text below no longer reflects — this
' function has routed through Explorer-first since Phase 4e. Firm policy as
' of 2026-08-28 also now assumes the desktop client is installed on every
' workstation, so the Explorer branch should be the common case, not the
' fallback. Original paragraph kept below for history; do not read it as a
' description of current behavior.
'
' Pre-check existence with GetMetadata: if the folder doesn't exist, offer to
' create it (restores the legacy S:\ "Folder doesn't exist. Do you want to
' create it?" behavior — the cmdCreateFolder / cmdCreateFolderSub buttons on
' frmClientLedger route here, as do all the plain open-folder buttons).
' Creation goes through DropboxService.CreateFolder against the SP-resolved
' path (G27).
'
' Create is a write op: gated by ALLOW_DROPBOX_WRITES. In the test build it
' raises vbObjectError + 6001, intercepted in Err_Handler with the friendly
' "writes disabled" message; flip the kill-switch to exercise it for real.
'
' Original (now-stale) design decision, kept for history: do NOT route
' through Windows Explorer / the local-synced folder. The Dropbox desktop
' client is not a deployment prerequisite. Users accept the cosmetic
' "Unsupported path provided" banner Dropbox shows for /work/ team-namespace
' deep links.
Public Function OpenDocumentFolder(ByVal CaseID As Variant, ByVal DocumentType As Variant) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim FolderName As String
Dim pathForCheck As String
Dim found As Boolean
Dim errDetail As String
Dim mdJson As String
Dim apiOk As Boolean
Dim doOpen As Boolean

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
            pathForCheck = FolderName
            If Right$(pathForCheck, 1) = "/" Then
                pathForCheck = Left$(pathForCheck, Len(pathForCheck) - 1)
            End If

            apiOk = DropboxService.GetMetadata(pathForCheck, found, errDetail, mdJson)

            ' doOpen drives the single folder-open site below. Default True so
            ' "folder exists" and "GetMetadata transport failure" both open the
            ' web URL (we do not blind-create when existence is unknown).
            doOpen = True

            If apiOk And Not found Then
                ' Folder confirmed missing — restore the legacy create-on-demand-
                ' with-confirmation flow, mapped to Dropbox via CreateFolder.
                ' pathForCheck is SP-resolved (never built in VBA).
                If MsgBox("This folder doesn't exist yet:" & vbCrLf & vbCrLf & _
                          FolderName & vbCrLf & vbCrLf & _
                          "Do you want to create it now?", _
                          vbYesNo + vbQuestion, "TB CMS") = vbYes Then
                    If Not DropboxService.CreateFolder(pathForCheck) Then
                        MsgBox "The folder could not be created. Please contact " & _
                               "IT — see tblDropboxLog for detail.", _
                               vbExclamation, "TB CMS"
                        doOpen = False
                    End If
                Else
                    doOpen = False   ' user declined — nothing to open
                End If
            End If

            If doOpen Then
                ' Phase 4e (G24): prefer Windows Explorer (the firm's preferred
                ' view), fall back to the Dropbox web URL. Both args are SP-
                ' resolved — pathForCheck (trailing "/" stripped) drives the
                ' local mapping + on-disk probe; FolderName keeps the prior web
                ' deep-link form. VBA never builds a path; it only maps one.
                OpenFolderInExplorerOrWeb pathForCheck, FolderName

                ' LEGACY (Phase 4a/4b web-only open — pre-Phase 4e). To roll
                ' back to web-only, comment the call above and uncomment:
                '   webUrl = "https://www.dropbox.com/work" & FolderName
                '   Shell "explorer.exe """ & webUrl & """", vbNormalFocus
                ' (explorer.exe opens the URL in the default browser; it is the
                ' running shell process — no admin, not routed via cmd.exe, so
                ' '&' in the URL is safe inside the quotes.)
            End If
        End If
    End If
    rv = True
Exit_Handler:
    OpenDocumentFolder = rv
    Exit Function
Err_Handler:
    If Err.Number = vbObjectError + 6001 Then
        ' GuardWritesEnabled — test environment, writes disabled. The folder
        ' open itself is read-only; this fires only on the create branch.
        ' Return True so the caller does NOT also pop its generic
        ' "Failed to open folder..." box — we've already shown a clear message.
        MsgBox "Dropbox writes are disabled in this test build " & _
               "(ALLOW_DROPBOX_WRITES = False). The folder was not created.", _
               vbInformation, "TB CMS — test environment"
        rv = True
    Else
        rv = False
        foo = pcaStdErrMsg(Err, Error)
    End If
    Resume Exit_Handler
End Function


' ============================================================================
' OpenFolderInExplorerOrWeb   (Phase 4e — Explorer-first folder open; G24/G27)
' ----------------------------------------------------------------------------
' Opens an SP-resolved Dropbox folder, preferring Windows Explorer (the firm's
' preferred, more familiar environment) and falling back to the Dropbox web UI
' when the desktop client is absent or the folder has not synced locally yet.
'
'   dropboxPath - SP-resolved /Company/... path, trailing "/" stripped
'                 (used for the local mapping + the on-disk existence probe)
'   webPathRaw  - SP-resolved path as-is (used to build the web deep-link;
'                 preserves the existing trailing-slash web behavior)
'
' Team-space layout (important): Dropbox's info.json reports the MEMBER folder
' (e.g. ...\Tate Bywater Dropbox\Indra Sugianto), but the bridge addresses
' everything through the TEAM namespace, and on the desktop client team folders
' (/Company, ...) are mounted at the TEAM-SPACE root (...\Tate Bywater Dropbox\
' Company) — i.e. the PARENT of the member folder, not under it. So we try the
' team-space root first, then the member root, and open whichever exists.
'
' Decision order:
'   1. Resolve the desktop-client (member) root via DropboxService; lazy-resolve
'      once if not known this session.
'   2. Pick the sync root under which the top-level team folder (e.g. \Company)
'      actually exists on disk — the team-space root (parent of the member
'      folder) first, then the member root. (Team folders live at the team-space
'      root, NOT under the member folder.)
'   3. Open the requested folder in Explorer. If it has not synced down yet,
'      walk up to the NEAREST already-synced ancestor (never above the team
'      folder) and open that instead — so a large initial sync still lands the
'      user in Explorer near the target rather than bouncing to the browser.
'   4. If no sync root is found (no client, or the team folder not synced at
'      all), fall back to the Dropbox web URL (prior behavior).
'
' Existence is probed with GetAttr (robust for Dropbox online-only placeholder
' folders). Default-safe: with no client / nothing synced the routine behaves
' exactly as the pre-4e web-only open. VBA never builds a Dropbox path; both
' inputs are stored-proc output — this only maps an existing path to its local
' mount and shells the OS.
' ----------------------------------------------------------------------------
Private Sub OpenFolderInExplorerOrWeb(ByVal dropboxPath As String, _
                                      ByVal webPathRaw As String)
    Const CALLER As String = "OpenFolderInExplorerOrWeb"
    Dim memberRoot As String
    Dim parentRoot As String
    Dim rel As String
    Dim firstSeg As String
    Dim topFolder As String
    Dim syncRoot As String
    Dim full As String
    Dim best As String
    Dim pos As Long
    Dim webUrl As String

    On Error Resume Next        ' any mapping/probe/shell hiccup => web fallback

    ' 1. Resolve the Dropbox desktop-client (member) root.
    memberRoot = DropboxService.GetLocalSyncedRoot()
    If LenB(memberRoot) = 0 Then
        If DropboxService.ResolveLocalSyncedRoot() Then _
            memberRoot = DropboxService.GetLocalSyncedRoot()
    End If

    If LenB(memberRoot) > 0 And Left$(dropboxPath, 1) = "/" Then
        rel = Replace(dropboxPath, "/", "\")            ' "/Company/X" -> "\Company\X"
        pos = InStr(2, rel, "\")
        If pos > 0 Then firstSeg = Left$(rel, pos - 1) Else firstSeg = rel   ' "\Company"

        ' 2. Pick the root under which the top team folder (\Company) exists:
        '    team-space root (parent of member) first, then the member root.
        If InStrRev(memberRoot, "\") > 0 Then
            parentRoot = Left$(memberRoot, InStrRev(memberRoot, "\") - 1)
        End If
        If LenB(parentRoot) > 0 Then
            If LocalFolderExists(parentRoot & firstSeg) Then syncRoot = parentRoot
        End If
        If LenB(syncRoot) = 0 Then
            If LocalFolderExists(memberRoot & firstSeg) Then syncRoot = memberRoot
        End If

        If LenB(syncRoot) > 0 Then
            topFolder = syncRoot & firstSeg             ' exists; never go above this
            full = syncRoot & rel

            ' 3. Open the exact folder, else the nearest synced ancestor.
            best = full
            Do While Len(best) > Len(topFolder)
                If LocalFolderExists(best) Then Exit Do
                pos = InStrRev(best, "\")
                If pos <= 0 Then Exit Do
                best = Left$(best, pos - 1)
            Loop
            If Not LocalFolderExists(best) Then best = topFolder

            If StrComp(best, full, vbTextCompare) = 0 Then
                DropboxService.LogLocal CALLER, "Info", "Explorer open: " & best
            Else
                DropboxService.LogLocal CALLER, "Info", _
                    "Explorer open (nearest synced ancestor): " & best & " for " & full
                ' Don't open a parent silently — the exact folder isn't on this
                ' PC yet (still syncing, or it was deleted and the server view
                ' lags). Tell the user which folder we're actually opening.
                MsgBox "This case's exact folder isn't on your PC yet — it may " & _
                       "still be syncing from Dropbox." & vbCrLf & vbCrLf & _
                       "Opening the closest available folder in Windows Explorer:" & _
                       vbCrLf & "    " & best, _
                       vbInformation, "TB CMS"
            End If
            Shell "explorer.exe """ & best & """", vbNormalFocus
            Exit Sub
        End If
    End If

    ' 4. Fall back to the Dropbox web UI (client absent, or the team folder is
    '    not synced down at all yet).
    webUrl = "https://www.dropbox.com/work" & webPathRaw
    DropboxService.LogLocal CALLER, "Info", "Explorer unavailable; web open: " & webUrl
    Shell "explorer.exe """ & webUrl & """", vbNormalFocus
End Sub


' Robust local-folder existence probe. Uses GetAttr rather than Dir because it
' reads the directory entry without hydrating Dropbox online-only placeholders
' and is reliable for cloud-backed folders. Returns False on any error
' (not found / bad path).
Private Function LocalFolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    Dim attrs As Long
    attrs = GetAttr(p)
    If Err.Number = 0 Then LocalFolderExists = ((attrs And vbDirectory) = vbDirectory)
    Err.Clear
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
' the file to %TEMP%\TBCMS\<GUID>_<filename> and opens the local path via
' explorer.exe (see G27) so the document opens in its native app.
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


' ============================================================================
' PHASE 5 — END-TO-END HAPPY-PATH WORKFLOW TEST
' ============================================================================
' Self-contained sequenced test of all three Phase 4d write workflows on a
' designated test case, with strict cleanup between steps:
'
'   Step 1  CopyDocumentToClosedFileScan
'           - Pre-stages: uploads a dummy file into the open folder so the
'             source actually exists in Dropbox (cases with 0 documents have
'             no folder in Dropbox yet)
'           - Calls CopyDocumentToClosedFileScan
'           - Verifies: destination folder appears under
'             /Company/Closed File Scans/TB/<yr>/<caseFolder>
'           - Cleans up: deletes the copy + the staged open folder
'
'   Step 2  SaveScannedFileAs (semi-interactive — user clicks SaveAs dialog)
'           - Creates a local temp file
'           - Calls SaveScannedFileAs(CaseID, "Case Notes", local, "Open")
'           - Verifies: file lands in the open folder + a tblCaseDocuments
'             row was inserted
'           - Cleans up: deletes file + row + the open folder
'
'   Step 3  MoveDocumentByCaseStatus (forward + reverse)
'           - Pre-stages: uploads a stub file + INSERTs a matching
'             tblCaseDocuments row so the G2 SP has 1 row to update
'             (so we hit the exactly-one-zero "Success" branch instead of
'             the both-zero throw branch)
'           - Calls MoveDocumentByCaseStatus(CaseID, "Closed")
'           - Verifies: folder relocated to _CLOSED, row path updated
'           - Calls MoveDocumentByCaseStatus(CaseID, "Open")
'           - Verifies: folder back, row reverted
'           - Cleans up: deletes stub file + row + the open folder
'
' Requires:
'   - DropboxService.ALLOW_DROPBOX_WRITES = True
'   - A valid Dropbox token
'   - testCaseID must be an Open case with 0 rows in tblCaseDocuments AND
'     tblScans (pre-flight enforces the tblCaseDocuments check)
'
' Returns a multi-line summary. On failure, attempts best-effort cleanup
' (including reversing a Move that already went to Closed) before returning.
'
' Usage from VBA Immediate window:
'   ? DocumentManagement.Phase5_E2E_HappyPathTest(30405)
Public Function Phase5_E2E_HappyPathTest(ByVal testCaseID As Long) As String
    Const STUB_DOCTYPE As String = "Case Notes"
    Const CALLER_TAG As String = "Phase5-Test"

    Dim resultLines As String
    Dim stepName As String

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fso As Object
    Dim ticker As String
    Dim found As Boolean
    Dim errDetail As String
    Dim mdJson As String
    Dim actualPath As String

    Dim openFolder As String
    Dim closedFolder As String
    Dim closedScanFolder As String
    Dim caseFolderBase As String

    ' Step 1 state
    Dim step1_OpenFolderStaged As Boolean
    Dim step1_StagedFile As String
    Dim step1_CopyDest As String
    Dim step1_CopyMade As Boolean

    ' Step 2 state
    Dim step2_LocalTemp As String
    Dim step2_UploadedPath As String
    Dim step2_DocID As Long

    ' Step 3 state
    Dim step3_StubFile As String
    Dim step3_StubRowID As Long
    Dim step3_FolderInClosed As Boolean

    On Error GoTo HandleError
    resultLines = "Phase5_E2E_HappyPathTest CaseID=" & testCaseID & vbCrLf

    ' --- 0. Pre-flight ----------------------------------------------------
    stepName = "0.PreFlight.WritesGate"
    If Not DropboxService.ALLOW_DROPBOX_WRITES Then
        Phase5_E2E_HappyPathTest = _
            "SKIP: ALLOW_DROPBOX_WRITES = False. Flip the constant in " & _
            "DropboxService.bas, re-import, then re-run."
        Exit Function
    End If

    stepName = "0.PreFlight.ConfigLoad"
    DropboxService.InitializeDropboxConfig

    stepName = "0.PreFlight.EnsureValidToken"
    If Not DropboxService.EnsureValidToken() Then _
        Err.Raise vbObjectError + 7002, , _
            "No valid Dropbox token. Run Phase3b_Pass2_AuthFlowTest first."

    stepName = "0.PreFlight.ResolvePaths"
    openFolder = GetDocumentFolderName(testCaseID, "General")
    closedFolder = GetClosedDocumentFolderName(testCaseID, "General")
    closedScanFolder = GetClosedFileScanFolderName(testCaseID, "General")
    If pcaempty(openFolder) Or pcaempty(closedFolder) Or pcaempty(closedScanFolder) Then _
        Err.Raise vbObjectError + 7000, , _
            "Path resolution returned empty. open=[" & openFolder & _
            "] closed=[" & closedFolder & "] closedScan=[" & closedScanFolder & "]"

    If Right$(openFolder, 1) = "/" Then openFolder = Left$(openFolder, Len(openFolder) - 1)
    If Right$(closedFolder, 1) = "/" Then closedFolder = Left$(closedFolder, Len(closedFolder) - 1)
    If Right$(closedScanFolder, 1) = "/" Then closedScanFolder = Left$(closedScanFolder, Len(closedScanFolder) - 1)
    caseFolderBase = Mid$(openFolder, InStrRev(openFolder, "/") + 1)

    resultLines = resultLines & _
        "  open  : " & openFolder & vbCrLf & _
        "  closed: " & closedFolder & vbCrLf & _
        "  scan  : " & closedScanFolder & vbCrLf

    stepName = "0.PreFlight.StartingState"
    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString
    Set rs = cn.Execute("SELECT COUNT(*) AS N FROM dbo.tblCaseDocuments WHERE CaseID = " & testCaseID)
    If CLng(pcaConvertNulls(rs("N"), 0)) <> 0 Then _
        Err.Raise vbObjectError + 7001, , _
            "Pre-flight: tblCaseDocuments has " & rs("N") & " rows for CaseID " & _
            testCaseID & "; expected 0. Pick a different case."
    rs.Close

    ticker = Format$(Now, "yyyymmdd_hhnnss")
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    MkDir Environ$("TEMP") & "\TBCMS"
    Err.Clear
    On Error GoTo HandleError

    ' --- 1. CopyDocumentToClosedFileScan ----------------------------------
    stepName = "1.Setup.UploadStagedFile"
    Dim step1Local As String
    step1Local = Environ$("TEMP") & "\TBCMS\phase5_step1_stage_" & ticker & ".txt"
    With fso.CreateTextFile(step1Local, True)
        .Write "phase5 step1 stage at " & Now()
        .Close
    End With
    step1_StagedFile = openFolder & "/__phase5_step1_stage__.txt"
    If Not DropboxService.UploadFile(step1Local, step1_StagedFile, Null, CALLER_TAG) Then _
        Err.Raise vbObjectError + 7010, , _
            "UploadFile failed for step 1 stage; see tblDropboxAuditLog"
    step1_OpenFolderStaged = True
    On Error Resume Next
    fso.DeleteFile step1Local, True
    Err.Clear
    On Error GoTo HandleError

    stepName = "1.Run.CopyDocumentToClosedFileScan"
    If Not CopyDocumentToClosedFileScan(testCaseID) Then _
        Err.Raise vbObjectError + 7011, , "CopyDocumentToClosedFileScan returned False"
    step1_CopyDest = closedScanFolder & "/" & caseFolderBase
    step1_CopyMade = True

    stepName = "1.Verify.CopyDestExists"
    If Not DropboxService.GetMetadata(step1_CopyDest, found, errDetail, mdJson) Then _
        Err.Raise vbObjectError + 7012, , _
            "GetMetadata transport failure for copy dest: " & errDetail
    If Not found Then _
        Err.Raise vbObjectError + 7013, , "Copy destination not found: " & step1_CopyDest
    resultLines = resultLines & "Step 1 OK — copy at: " & step1_CopyDest & vbCrLf

    stepName = "1.Cleanup.DeleteCopy"
    If Not DropboxService.DeleteFile(step1_CopyDest, Null, CALLER_TAG) Then _
        Err.Raise vbObjectError + 7014, , "Failed to delete copy at " & step1_CopyDest
    step1_CopyMade = False

    stepName = "1.Cleanup.DeleteStagedFolder"
    If Not DropboxService.DeleteFile(openFolder, Null, CALLER_TAG) Then _
        Err.Raise vbObjectError + 7015, , "Failed to delete staged open folder " & openFolder
    step1_OpenFolderStaged = False
    step1_StagedFile = ""

    ' --- 2. SaveScannedFileAs (semi-interactive) --------------------------
    stepName = "2.Setup.LocalTempFile"
    step2_LocalTemp = Environ$("TEMP") & "\TBCMS\phase5_step2_source_" & ticker & ".txt"
    With fso.CreateTextFile(step2_LocalTemp, True)
        .Write "phase5 step2 source content at " & Now()
        .Close
    End With

    stepName = "2.Run.SaveScannedFileAs"
    resultLines = resultLines & _
        "Step 2 — invoking SaveScannedFileAs. You will see a SaveAs dialog; " & _
        "accept the suggested filename to proceed (cancel will abort the test)." & vbCrLf
    If Not SaveScannedFileAs(testCaseID, STUB_DOCTYPE, step2_LocalTemp, "Open") Then _
        Err.Raise vbObjectError + 7020, , "SaveScannedFileAs returned False"

    stepName = "2.Verify.tblCaseDocumentsRow"
    Set rs = cn.Execute( _
        "SELECT TOP 1 CaseDocumentID, DocumentFileName FROM dbo.tblCaseDocuments " & _
        "WHERE CaseID = " & testCaseID & _
        " AND DocumentType = " & pcaAddQuotes(STUB_DOCTYPE) & _
        " ORDER BY CaseDocumentID DESC")
    If rs.EOF Then _
        Err.Raise vbObjectError + 7021, , _
            "No tblCaseDocuments row found for CaseID " & testCaseID & _
            " (did you cancel the SaveAs dialog?)"
    step2_DocID = CLng(rs("CaseDocumentID"))
    step2_UploadedPath = pcaConvertNulls(rs("DocumentFileName"), "")
    rs.Close

    stepName = "2.Verify.UploadLanded"
    If Not DropboxService.GetMetadata(step2_UploadedPath, found, errDetail, mdJson) Then _
        Err.Raise vbObjectError + 7022, , _
            "GetMetadata transport failure for uploaded path: " & errDetail
    If Not found Then _
        Err.Raise vbObjectError + 7023, , _
            "Uploaded file not found at " & step2_UploadedPath
    resultLines = resultLines & _
        "Step 2 OK — uploaded to: " & step2_UploadedPath & _
        " (CaseDocumentID=" & step2_DocID & ")" & vbCrLf

    stepName = "2.Cleanup.DeleteFromDropbox"
    If Not DropboxService.DeleteFile(step2_UploadedPath, Null, CALLER_TAG) Then _
        Err.Raise vbObjectError + 7024, , _
            "Failed to delete uploaded file from Dropbox: " & step2_UploadedPath
    step2_UploadedPath = ""

    stepName = "2.Cleanup.DeleteOpenFolder"
    ' Best-effort: the open folder may already be empty/auto-removed.
    On Error Resume Next
    DropboxService.DeleteFile openFolder, Null, CALLER_TAG
    Err.Clear
    On Error GoTo HandleError

    stepName = "2.Cleanup.DeleteRow"
    cn.Execute "DELETE FROM dbo.tblCaseDocuments WHERE CaseDocumentID = " & step2_DocID
    step2_DocID = 0

    On Error Resume Next
    fso.DeleteFile step2_LocalTemp, True
    Err.Clear
    On Error GoTo HandleError
    step2_LocalTemp = ""

    ' --- 3. MoveDocumentByCaseStatus (forward + reverse) ------------------
    stepName = "3.Setup.UploadStubFile"
    Dim step3Local As String
    step3Local = Environ$("TEMP") & "\TBCMS\phase5_step3_stub_" & ticker & ".txt"
    With fso.CreateTextFile(step3Local, True)
        .Write "phase5 step3 stub at " & Now()
        .Close
    End With
    step3_StubFile = openFolder & "/__phase5_step3_stub__.txt"
    If Not DropboxService.UploadFile(step3Local, step3_StubFile, Null, CALLER_TAG) Then _
        Err.Raise vbObjectError + 7030, , "Stub UploadFile failed"
    On Error Resume Next
    fso.DeleteFile step3Local, True
    Err.Clear
    On Error GoTo HandleError

    stepName = "3.Setup.InsertStubRow"
    cn.Execute _
        "INSERT INTO dbo.tblCaseDocuments (CaseID, DocumentType, DocumentFileName, CreatedOn) " & _
        "VALUES (" & testCaseID & ", " & pcaAddQuotes(STUB_DOCTYPE) & ", " & _
        pcaAddQuotes(step3_StubFile) & ", GETDATE())"
    Set rs = cn.Execute("SELECT CAST(@@IDENTITY AS int) AS NewID")
    step3_StubRowID = CLng(rs("NewID"))
    rs.Close

    stepName = "3.Run.MoveToClosed"
    If Not MoveDocumentByCaseStatus(testCaseID, "Closed") Then _
        Err.Raise vbObjectError + 7031, , _
            "MoveDocumentByCaseStatus(Closed) returned False"
    step3_FolderInClosed = True

    stepName = "3.Verify.AfterMoveToClosed"
    Dim expectedClosedFile As String
    expectedClosedFile = Replace(step3_StubFile, openFolder, closedFolder)
    If Not DropboxService.GetMetadata(expectedClosedFile, found, errDetail, mdJson) Then _
        Err.Raise vbObjectError + 7032, , _
            "GetMetadata transport failure post-MoveToClosed: " & errDetail
    If Not found Then _
        Err.Raise vbObjectError + 7033, , _
            "Stub file not at expected closed path after move: " & expectedClosedFile

    Set rs = cn.Execute( _
        "SELECT DocumentFileName FROM dbo.tblCaseDocuments WHERE CaseDocumentID = " & step3_StubRowID)
    actualPath = pcaConvertNulls(rs("DocumentFileName"), "")
    rs.Close
    If actualPath <> expectedClosedFile Then _
        Err.Raise vbObjectError + 7034, , _
            "tblCaseDocuments path not updated to closed. Expected=[" & _
            expectedClosedFile & "] Got=[" & actualPath & "]"
    resultLines = resultLines & _
        "Step 3a OK — Move(Closed): folder + row at " & expectedClosedFile & vbCrLf

    stepName = "3.Run.MoveBackToOpen"
    If Not MoveDocumentByCaseStatus(testCaseID, "Open") Then _
        Err.Raise vbObjectError + 7035, , _
            "MoveDocumentByCaseStatus(Open) returned False"
    step3_FolderInClosed = False

    stepName = "3.Verify.AfterMoveToOpen"
    If Not DropboxService.GetMetadata(step3_StubFile, found, errDetail, mdJson) Then _
        Err.Raise vbObjectError + 7036, , _
            "GetMetadata transport failure post-MoveToOpen: " & errDetail
    If Not found Then _
        Err.Raise vbObjectError + 7037, , _
            "Stub file not at expected open path after revert: " & step3_StubFile

    Set rs = cn.Execute( _
        "SELECT DocumentFileName FROM dbo.tblCaseDocuments WHERE CaseDocumentID = " & step3_StubRowID)
    actualPath = pcaConvertNulls(rs("DocumentFileName"), "")
    rs.Close
    If actualPath <> step3_StubFile Then _
        Err.Raise vbObjectError + 7038, , _
            "tblCaseDocuments path not reverted to open. Expected=[" & _
            step3_StubFile & "] Got=[" & actualPath & "]"
    resultLines = resultLines & _
        "Step 3b OK — Move(Open): folder + row reverted to " & step3_StubFile & vbCrLf

    stepName = "3.Cleanup.DeleteStubRow"
    cn.Execute "DELETE FROM dbo.tblCaseDocuments WHERE CaseDocumentID = " & step3_StubRowID
    step3_StubRowID = 0

    stepName = "3.Cleanup.DeleteStubFolder"
    If Not DropboxService.DeleteFile(openFolder, Null, CALLER_TAG) Then _
        Err.Raise vbObjectError + 7039, , _
            "Failed to delete open folder " & openFolder
    step3_StubFile = ""

    cn.Close
    Set cn = Nothing

    resultLines = resultLines & vbCrLf & "OK — all 3 workflows passed."
    Phase5_E2E_HappyPathTest = resultLines
    Exit Function

HandleError:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next

    ' Step 3 cleanup
    If step3_FolderInClosed Then
        ' Forward move went through; try to put folder back to Open
        MoveDocumentByCaseStatus testCaseID, "Open"
    End If
    If step3_StubRowID > 0 And Not cn Is Nothing Then
        cn.Execute "DELETE FROM dbo.tblCaseDocuments WHERE CaseDocumentID = " & step3_StubRowID
    End If
    If LenB(step3_StubFile) > 0 Then _
        DropboxService.DeleteFile step3_StubFile, Null, CALLER_TAG & "-Cleanup"

    ' Step 2 cleanup
    If LenB(step2_UploadedPath) > 0 Then _
        DropboxService.DeleteFile step2_UploadedPath, Null, CALLER_TAG & "-Cleanup"
    If step2_DocID > 0 And Not cn Is Nothing Then
        cn.Execute "DELETE FROM dbo.tblCaseDocuments WHERE CaseDocumentID = " & step2_DocID
    End If
    If LenB(step2_LocalTemp) > 0 And Not fso Is Nothing Then _
        fso.DeleteFile step2_LocalTemp, True

    ' Step 1 cleanup
    If step1_CopyMade And LenB(step1_CopyDest) > 0 Then _
        DropboxService.DeleteFile step1_CopyDest, Null, CALLER_TAG & "-Cleanup"
    If step1_OpenFolderStaged And LenB(openFolder) > 0 Then _
        DropboxService.DeleteFile openFolder, Null, CALLER_TAG & "-Cleanup"

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If

    Phase5_E2E_HappyPathTest = resultLines & vbCrLf & _
        "FAIL at " & stepName & ": Err=" & errNum & " " & errDesc
End Function
