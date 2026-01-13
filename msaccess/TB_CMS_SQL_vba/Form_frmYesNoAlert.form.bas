' Component: Form_frmYesNoAlert
' Type: document
' Lines: 25
' ============================================================

Option Compare Database

Private Sub cmdNo_Click()
    YesNo_Value = 2
    Util.CloseForm
End Sub

Private Sub cmdYes_Click()
    YesNo_Value = 1
    Util.CloseForm
End Sub

Public Function CodeBehind_Form_Load() As Boolean
    YesNo_Value = -1
    
    lblMessage.Caption = message
    lblMessage2.Caption = Message1
    
    CodeBehind_Form_Load = True
End Function


Private Sub Form_Load()
CodeBehind_Form_Load
End Sub