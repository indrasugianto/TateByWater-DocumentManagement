' Component: Form_frmBrowse
' Type: document
' Lines: 38
' ============================================================

Option Compare Database

Dim args As String

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
    args = Nz(Me.OpenArgs, "")
End Function

Private Sub Command6_Click()
    Dim dbName As String
    
    Dim wrack As String
    args = "All Excel Files--*.xls,*.xlsx"
    If args <> "" Then
        
        Dim arr() As String
        arr = Split(args, "--")
        
        dbName = BrowseDB("Select path for backend DB", arr(0), arr(1))
    Else
        dbName = BrowseDB("Select path for backend DB")
    End If
    
    lblMessage.Caption = dbName
End Sub