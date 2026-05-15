Attribute VB_Name = "DocumentManagement"
Option Compare Database
Option Explicit
Dim foo

' =============================================================================
' Phase 4 rewire status (last updated 2026-05-15):
'
'   Phase 4a — Read-flow rewire (DONE in this file):
'     OpenDocumentFile     -> DropboxService.OpenDocument
'                             (download to %TEMP%\TBCMS\ + native-app launch)
'     OpenDocumentFolder   -> Dropbox web URL (Application.FollowHyperlink)
'
'   Rollback safety: the legacy (pre-4a) bodies of OpenDocumentFile and
'   OpenDocumentFolder are preserved as commented-out blocks immediately
'   above each rewired function. To roll back, swap the comment/uncomment
'   state on the active vs. LEGACY block and re-import this file. The
'   exact swap procedure is documented above the first LEGACY block.
'
'   Phase 4b / 4c / 4d — pending:
'     4b config layer (GetDocumentRootFolder/GetScannerFolder switch to
'        tblDropboxRootConfig; add LocalPathToDropboxPath to DropboxService),
'     4c G2/G13 SQL changes,
'     4d write-flow rewires (SaveScannedFileAs / MoveDocumentByCaseStatus /
'        CopyDocumentToClosedFileScan). Until 4d lands, those functions still
'        contain the legacy S:\ FSO/FileCopy logic; calling them in the test
'        environment will FAIL because the SQL-stored paths are now /Company/-
'        rooted (not valid Windows paths). This is intentional — write flows
'        are gated by ALLOW_DROPBOX_WRITES = False in DropboxService anyway,
'        and 4a delivers the read-only validation surface for the Phase 6.5
'        50-row open-document gate.
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


' Phase 4b will switch this from tblDocumentRootDirectory to
' tblDropboxRootConfig.TeamRootPath. Unchanged in 4a.
Public Function GetDocumentRootFolder() As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "SELECT DocumentRootDirectory "
    sql = sql & "FROM tblDocumentRootDirectory"

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = rs("DocumentRootDirectory")
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


' Phase 4b will switch this from tblDocumentRootDirectory to
' tblDropboxRootConfig.ScannerDirectory. Unchanged in 4a.
Public Function GetScannerFolder() As String
On Error GoTo Err_Handler
Dim rv As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString

    sql = ""
    sql = sql & "SELECT ScannerDirectory "
    sql = sql & "FROM tblDocumentRootDirectory"

    Set rs = cn.Execute(sql)

    If Not rs.EOF() Then
        rv = rs("ScannerDirectory")
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


' --- Phase 4d (pending) — legacy write flow ---------------------------------
' Currently still uses S:\ FileCopy + spSaveCaseDocument. Will be rewritten
' in 4d to use DropboxService.UploadFile / UploadLargeFile + the orphan-file
' compensation policy. Until then, calling this in the test environment will
' fail because the SQL-stored DocumentFileName values are /Company/-rooted.
Public Function SaveScannedFileAs(ByVal CaseID As Integer, ByVal DocumentType As String, ByVal SourceFileName As String, ByVal CaseStatus As String) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim FolderName As String
Dim DestinationFileName As String

Dim fDialog As Office.FileDialog
Dim varFile As Variant

    If CaseStatus = "Closed" Then
        FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
    Else
        FolderName = GetDocumentFolderName(CaseID, DocumentType)
    End If

    DestinationFileName = GetDocumentFileName(CaseID, DocumentType)
    DestinationFileName = DestinationFileName & "." & Right(SourceFileName, Len(SourceFileName) - InStrRev(SourceFileName, "."))

    If FolderExistsCreate(FolderName, True) Then
        Set fDialog = Application.FileDialog(msoFileDialogSaveAs)

        With fDialog
            .AllowMultiSelect = False
            .InitialFileName = FolderName
            .InitialFileName = FolderName & DestinationFileName

            If .show = True Then
                For Each varFile In .SelectedItems
                    FileCopy SourceFileName, varFile

                    If DocumentType = "Closed Final" Then
                        If MsgBox("Do you want to save the file in Closed File Scans directory?", vbYesNo, "TB CMS") = vbYes Then
                            FolderName = GetClosedFileScanFolderName(CaseID, "General")

                            If FolderExistsCreate(FolderName, True) Then
                                FileCopy SourceFileName, FolderName & DestinationFileName
                            End If
                        End If
                    End If

                    If Not SaveCaseDocument(CaseID, DocumentType, varFile) Then
                        MsgBox "Fail to save case document record...", , "TB CMS"
                    End If
                Next
            End If
        End With
    End If

    rv = True
Exit_Handler:
    SaveScannedFileAs = rv
    Exit Function
Err_Handler:
    rv = False
    Resume Exit_Handler
End Function


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


' --- Phase 4a rewire --------------------------------------------------------
' Folder open: route through the Dropbox web URL. The previous behavior
' resolved an S:\ folder, prompted to create it if missing, and showed a
' local file-picker dialog so the user could pick a file inside. In the
' Dropbox world, the folder lives on Dropbox; opening it in the browser
' gives the user access to navigate, click files, drag-and-drop, etc.,
' via Dropbox's native UI. The desktop client (if installed and signed in)
' also exposes the folder locally, but the web URL works regardless of
' client state.
'
' UX TODO (Phase 6.5 UAT): validate that browser-open is the right UX vs.
' opening the local-synced folder via Explorer. See plan Phase 5 step 1.
Public Function OpenDocumentFolder(ByVal CaseID As Variant, ByVal DocumentType As Variant) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim FolderName As String
Dim webUrl As String
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
            ' Pre-check: confirm the folder exists in Dropbox before opening
            ' the web URL. Otherwise Dropbox shows its own "Unsupported path
            ' provided" message which is confusing — better to surface a
            ' clear "folder doesn't exist yet" here. GetMetadata expects no
            ' trailing slash on folder paths.
            pathForCheck = FolderName
            If Right$(pathForCheck, 1) = "/" Then
                pathForCheck = Left$(pathForCheck, Len(pathForCheck) - 1)
            End If

            apiOk = DropboxService.GetMetadata(pathForCheck, found, errDetail, mdJson)

            If apiOk And Not found Then
                ' Folder definitively does not exist (HTTP 200 + path/not_found)
                MsgBox "The Dropbox folder for this case doesn't exist yet:" & vbCrLf & vbCrLf & _
                       FolderName & vbCrLf & vbCrLf & _
                       "This typically means no documents have been saved for this " & _
                       "case + document-type combination yet. The folder will be " & _
                       "created automatically when the first document is saved.", _
                       vbExclamation, "TB CMS"
            Else
                ' Either GetMetadata confirmed found=True, or there was a
                ' transport failure — proceed to the URL and let Dropbox respond.
                ' /work/ is the Dropbox Business team-content prefix; /home/
                ' triggers an "Unsupported path provided" warning banner for
                ' team-namespace content even though it functionally works.
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


' --- Phase 4d (pending) — legacy write flow ---------------------------------
' Currently still uses FSO.CopyFolder + FSO.DeleteFolder + legacy
' spMoveDocumentFolder. Will be rewritten in 4d to use
' DropboxService.MoveFile + new G2 spMoveDocumentFolder (@OldFolderPath /
' @NewFolderPath signature) + Dropbox-rollback compensation if the SP
' returns both-zero rowcount. Until then, calling this against /Company/-
' rooted SQL paths will fail at FSO.FolderExists("/Company/...").
Public Function MoveDocumentByCaseStatus(ByVal CaseID As Variant, ByVal CaseStatus As String) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim SourceFolder As String
Dim TargetFolder As String
Dim FSO As Object
Dim cn As ADODB.Connection
Dim sql As String
Dim i As Integer
Dim LArray() As String

    rv = False
    If CaseStatus = "Closed" Then
        SourceFolder = GetDocumentFolderName(CaseID, "General")
        TargetFolder = GetClosedDocumentFolderName(CaseID, "General")
    Else
        SourceFolder = GetClosedDocumentFolderName(CaseID, "Init Intake, Notes, Documents")
        TargetFolder = GetDocumentFolderName(CaseID, "Init Intake, Notes, Documents")
    End If

    LArray = Split(TargetFolder, "\")
    i = 0
    Do While (LArray(i) <> "")
            i = i + 1
    Loop
    TargetFolder = Left(TargetFolder, Len(TargetFolder) - Len(LArray(i - 1)) - 1)

    If Right(SourceFolder, 1) = "\" Then
        SourceFolder = Left(SourceFolder, Len(SourceFolder) - 1)
    End If

    If Right(TargetFolder, 1) = "\" Then
        TargetFolder = Left(TargetFolder, Len(TargetFolder) - 1)
    End If

    Set FSO = CreateObject("scripting.filesystemobject")

    If Not FSO.FolderExists(SourceFolder) Then
        MsgBox "Source folder doesn't exists...", , "TB CMS"
        rv = False
    Else
        If Not FolderExistsCreate(TargetFolder, True) Then
            MsgBox "Failed to create target folder", , "TB CMS"
        Else
            FSO.CopyFolder Source:=SourceFolder, Destination:=TargetFolder
            FSO.DeleteFolder SourceFolder

            Set cn = New ADODB.Connection
            cn.Open PcaGetConnnectionString

            sql = ""
            sql = sql & "exec spMoveDocumentFolder "
            sql = sql & "@CaseID = " & CaseID
            sql = sql & ",@CaseStatus = " & pcaAddQuotes(CaseStatus)

            cn.Execute sql
            rv = True
        End If
    End If
Exit_Handler:
    MoveDocumentByCaseStatus = rv
    Exit Function
Err_Handler:
    If Err = 70 Then
        Call MsgBox("The application was not able to delete the original folder after copying it to the target folder.  Please manually delete this folder: " & vbCrLf & SourceFolder, vbExclamation, "TB CMS")
        Resume Next
    Else
        foo = pcaStdErrMsg(Err, Error)
    End If
    rv = False
    Resume Exit_Handler
End Function


' --- Phase 4d (pending) — legacy write flow ---------------------------------
' Currently uses FSO.CopyFolder. Will be rewritten in 4d to call
' DropboxService.CopyFile (folder-level — Dropbox copy_v2 handles both files
' and folders) with autorename=false per plan conflict policy.
Public Function CopyDocumentToClosedFileScan(ByVal CaseID As Variant) As Boolean
On Error GoTo Err_Handler
Dim rv As Boolean
Dim SourceFolder As String
Dim TargetFolder As String
Dim FSO As Object
Dim i As Integer
Dim LArray() As String

    rv = False
    SourceFolder = GetDocumentFolderName(CaseID, "General")
    TargetFolder = GetClosedFileScanFolderName(CaseID, "General")

    LArray = Split(TargetFolder, "\")
    i = 0
    Do While (LArray(i) <> "")
            i = i + 1
    Loop
    TargetFolder = Left(TargetFolder, Len(TargetFolder) - Len(LArray(i - 1)) - 1)

    If Right(SourceFolder, 1) = "\" Then
        SourceFolder = Left(SourceFolder, Len(SourceFolder) - 1)
    End If

    If Right(TargetFolder, 1) = "\" Then
        TargetFolder = Left(TargetFolder, Len(TargetFolder) - 1)
    End If

    Set FSO = CreateObject("scripting.filesystemobject")

    If Not FSO.FolderExists(SourceFolder) Then
        MsgBox "Source folder doesn't exists...", , "TB CMS"
        rv = False
    Else
        If Not FolderExistsCreate(TargetFolder, True) Then
            MsgBox "Fail to create target folder", , "TB CMS"
        Else
            FSO.CopyFolder Source:=SourceFolder, Destination:=TargetFolder
            rv = True
        End If
    End If
Exit_Handler:
    CopyDocumentToClosedFileScan = rv
    Exit Function
Err_Handler:
    foo = pcaStdErrMsg(Err, Error)
    rv = False
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
