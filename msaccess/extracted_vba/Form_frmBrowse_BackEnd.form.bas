' Component: Form_frmBrowse_BackEnd
' Type: document
' Lines: 46
' ============================================================

Option Compare Database

Private Sub cmdNo_Click()
    YesNo_Value = 2
    Util.CloseForm
End Sub

Private Sub cmdYes_Click()
    SelectedPath = lblMessage.Caption
    YesNo_Value = 1
    Util.CloseForm
End Sub

Public Function CodeBehind_Form_Load() As Boolean
    YesNo_Value = -1
    CodeBehind_Form_Load = True
End Function

Private Sub Command6_Click()
    Dim dbName As String
    dbName = BrowseDB_1("Select path for backend DB")
    lblMessage.Caption = dbName
End Sub

Private Sub Form_Load()
CodeBehind_Form_Load
End Sub


Public Function BrowseDB_1(Optional Title As String = "Select Access DB", Optional Description As String = "Access Database", _
                            Optional Filter As String = "*.accdb") As String

    Dim d As FileDialog

    Set d = Application.FileDialog(msoFileDialogFilePicker)
    
    d.AllowMultiSelect = False
    d.Filters.Clear
    d.Title = Title
    d.Filters.Add Description, Filter
    d.show
    On Error Resume Next
    BrowseDB_1 = d.SelectedItems(1)
    
End Function
