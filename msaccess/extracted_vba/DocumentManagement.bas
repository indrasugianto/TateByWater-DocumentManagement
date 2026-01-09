' Component: DocumentManagement
' Type: module
' Lines: 841
' ============================================================

Option Compare Database
Option Explicit
Dim foo



Public Function GetDocumentFileName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetDocumentFileName = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function


Public Function GetDocumentFolderName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetDocumentFolderName = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function


Public Function GetIntakeFolderName() As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetIntakeFolderName = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function

Public Function GetClosedDocumentFolderName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetClosedDocumentFolderName = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function


Public Function GetDocumentRootFolder() As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetDocumentRootFolder = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function

Public Function GetScannerFolder() As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetScannerFolder = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function

Public Function GetClosedFileScanFolderName(ByVal CaseID As Long, ByVal DocumentType As String) As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetClosedFileScanFolderName = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function


Public Function GetAllInvoicesFolderName(ByVal CaseID As Long) As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetAllInvoicesFolderName = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function


Function FolderExistsCreate(DirectoryPath As String, CreateIfNot As Boolean) As Boolean
    On Error GoTo ERR_HANDLER
    Dim rv As Boolean
    Dim elm As Variant
    Dim strCheckPath As String
    If Right(DirectoryPath, 1) <> "\" Then
        DirectoryPath = DirectoryPath & "\"
    End If
    
    If Dir(DirectoryPath, vbDirectory) <> "" Then
        rv = True
    Else
        ' Doesn't Exist Determine If user Wants to create
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
EXIT_HANDLER:
    FolderExistsCreate = rv
    Exit Function
ERR_HANDLER:
    rv = False
    Resume EXIT_HANDLER
End Function

Public Function OpenFileDialog(ByVal DialogBoxTitle As String, ByVal StartingFolder As String, ByVal FileExtention As String)
  
   ' Requires reference to Microsoft Office 11.0 Object Library.
 
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
 
 
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogOpen)
 
   With fDialog
 
      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
             
      ' Set the title of the dialog box.
      .Title = DialogBoxTitle
      
      'set the starting folder
      .InitialFileName = StartingFolder

 
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "All Files", "*.*"
 
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .show = True Then
         'Loop through each file selected and add it to our list box.
         For Each varFile In .SelectedItems
            Application.FollowHyperlink varFile
         Next
      End If
   End With
End Function

Public Function SelectFileDialog(ByVal DialogBoxTitle As String, ByVal StartingFolder As String, ByVal FileExtention As String) As String
  
   ' Requires reference to Microsoft Office 11.0 Object Library.
 
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant
   Dim rv As String
 
 
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
 
   With fDialog
 
      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
             
      ' Set the title of the dialog box.
      .Title = DialogBoxTitle
      
      'set the starting folder
      .InitialFileName = StartingFolder

 
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "All Files", "*.*"
 
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .show = True Then
         'Loop through each file selected and add it to our list box.
         For Each varFile In .SelectedItems
            rv = varFile
         Next
      Else
        rv = ""
      End If
   End With
   SelectFileDialog = rv
End Function

Public Function SaveScannedFileAs(ByVal CaseID As Integer, ByVal DocumentType As String, ByVal SourceFileName As String, ByVal CaseStatus As String) As Boolean
On Error GoTo ERR_HANDLER
Dim rv As Boolean
Dim FolderName As String
Dim DestinationFileName As String

Dim fDialog As Office.FileDialog
Dim varFile As Variant
Dim FileName As String


    If CaseStatus = "Closed" Then
        FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
    Else
        FolderName = GetDocumentFolderName(CaseID, DocumentType)
    End If
    
    DestinationFileName = GetDocumentFileName(CaseID, DocumentType)

    'append file extension to the destination file
    DestinationFileName = DestinationFileName & "." & Right(SourceFileName, Len(SourceFileName) - InStrRev(SourceFileName, "."))
     
    
    'check if the folder exists, if not create one
    If FolderExistsCreate(FolderName, True) Then
        Set fDialog = Application.FileDialog(msoFileDialogSaveAs)
        
        With fDialog
            .AllowMultiSelect = False
            .InitialFileName = FolderName
            
            .InitialFileName = FolderName & DestinationFileName
            
            
            If .show = True Then
                For Each varFile In .SelectedItems
                    'save the file
                    FileCopy SourceFileName, varFile
                    
                    'if it's closed final, put a copy under CLOSED FINAL SCAN folder as well
                    If DocumentType = "Closed Final" Then
                        If MsgBox("Do you want to save the file in Closed File Scans directory?", vbYesNo, "TB CMS") = vbYes Then
                            FolderName = GetClosedFileScanFolderName(CaseID, "General")
                        
                            If FolderExistsCreate(FolderName, True) Then
                                FileCopy SourceFileName, FolderName & DestinationFileName
                            End If
                        End If
                    End If
                    
                    'save the record
                    If Not SaveCaseDocument(CaseID, DocumentType, varFile) Then
                        MsgBox "Fail to save case document record...", , "TB CMS"
                    End If
                Next
            End If
        End With
    End If
    

    rv = True
EXIT_HANDLER:
    SaveScannedFileAs = rv
    Exit Function
ERR_HANDLER:
    rv = False
    Resume EXIT_HANDLER
End Function


Public Function SaveCaseDocument(ByVal CaseID As Integer, ByVal DocumentType As String, ByVal DocumentFileName As String) As Boolean
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    SaveCaseDocument = rv
    Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = False
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume

End Function

Public Function GetCaseDocument(ByVal CaseID As Integer, ByVal DocumentType) As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetCaseDocument = rv
    Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function
Public Function OpenDocumentFolder(ByVal CaseID As Variant, ByVal DocumentType) As Boolean
On Error GoTo ERR_HANDLER
Dim rv As Boolean
Dim FolderName As String

    rv = False
    
    If pcaempty(CaseID) Then
        MsgBox "Please select a case before proceeding...", , "TB CMS"
    Else
        If GetCaseClosedStatus(CaseID) Then
            FolderName = GetClosedDocumentFolderName(CaseID, DocumentType)
        Else
            FolderName = GetDocumentFolderName(CaseID, DocumentType)
        End If
        
        If Not FolderExistsCreate(FolderName, False) Then
            If MsgBox(FolderName & " Folder for this case doesn't exists.  Do you want to create it?", vbYesNo, "TB CMS") = vbYes Then
                If FolderExistsCreate(FolderName, True) Then
                    MsgBox "Document folder is created", , "TB CMS"
                    Call OpenFileDialog("Case Document", FolderName, "")
                End If
            End If
        Else
            Call OpenFileDialog("Case Document", FolderName, "")
        End If
    End If
    rv = True
EXIT_HANDLER:
    OpenDocumentFolder = rv
    Exit Function
ERR_HANDLER:
    rv = False
    Resume EXIT_HANDLER
End Function

Public Function OpenDocumentFile(ByVal CaseID As Variant, ByVal DocumentType As String) As Boolean
On Error GoTo ERR_HANDLER
Dim rv As Boolean
Dim DocumentFileName As String
    rv = False
    
    If pcaempty(CaseID) Then
        MsgBox "Please select a case before proceeding...", , "TB CMS"
    Else
        DocumentFileName = GetCaseDocument(CaseID, DocumentType)
        'check if the document exists
        If Not pcaempty(DocumentFileName) And Dir(DocumentFileName) <> "" Then
            Application.FollowHyperlink DocumentFileName
        Else
            MsgBox DocumentType & " is not found", vbExclamation, "TB CMS"
        End If
    End If
    rv = True
EXIT_HANDLER:
    OpenDocumentFile = rv
    Exit Function
ERR_HANDLER:
    rv = False
    Resume EXIT_HANDLER
End Function
Public Function MoveDocumentByCaseStatus(ByVal CaseID As Variant, ByVal CaseStatus As String) As Boolean
On Error GoTo ERR_HANDLER

    'move the client document to _CLOSED subfolder if CaseStatus = 'Closed'
    'move the client document back from _CLOSED subfolder if CaseStatus = 'Open'
    
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
       
    
    'move the TargetFolder one folder up
    LArray = Split(TargetFolder, "\")
    i = 0
    Do While (LArray(i) <> "")
            i = i + 1
    Loop
    TargetFolder = Left(TargetFolder, Len(TargetFolder) - Len(LArray(i - 1)) - 1)
    
    
    'remove "\" at the end
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
            
            'copy to the target folder
            FSO.CopyFolder Source:=SourceFolder, Destination:=TargetFolder
            
            FSO.DeleteFolder SourceFolder

            'update Case Document
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
EXIT_HANDLER:
    MoveDocumentByCaseStatus = rv
    Exit Function
ERR_HANDLER:
    If Err = 70 Then
        Call MsgBox("The application was not able to delete the original folder after copying it to the target folder.  Please manually delete this folder: " & vbCrLf & SourceFolder, vbExclamation, "TB CMS")
        Resume Next
    Else
        foo = pcaStdErrMsg(Err, Error)
    End If
    rv = False
    Resume EXIT_HANDLER
    Resume

End Function


Public Function CopyDocumentToClosedFileScan(ByVal CaseID As Variant) As Boolean
On Error GoTo ERR_HANDLER
   
Dim rv As Boolean
Dim SourceFolder As String
Dim TargetFolder As String
Dim FSO As Object
Dim cn As ADODB.Connection
Dim sql As String
Dim i As Integer
Dim LArray() As String



    rv = False
    SourceFolder = GetDocumentFolderName(CaseID, "General")
    TargetFolder = GetClosedFileScanFolderName(CaseID, "General")
       
    
    'move the TargetFolder one folder up
    LArray = Split(TargetFolder, "\")
    i = 0
    Do While (LArray(i) <> "")
            i = i + 1
    Loop
    TargetFolder = Left(TargetFolder, Len(TargetFolder) - Len(LArray(i - 1)) - 1)
    
    
    'remove "\" at the end
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
            
            'copy to the target folder
            FSO.CopyFolder Source:=SourceFolder, Destination:=TargetFolder
            
            rv = True
        End If
    End If
EXIT_HANDLER:
    CopyDocumentToClosedFileScan = rv
    Exit Function
ERR_HANDLER:
    foo = pcaStdErrMsg(Err, Error)
    rv = False
    Resume EXIT_HANDLER
    Resume
End Function


Public Function GetCaseClosedStatus(ByVal CaseID As Integer) As Boolean
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetCaseClosedStatus = rv
    Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = False
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume

End Function


Public Function GetIntakeDocumentFileName(ByVal IntakeID As Long) As String
On Error GoTo ERR_HANDLER
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
EXIT_HANDLER:
    GetIntakeDocumentFileName = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function
