' Component: Form_frmOkAlert
' Type: document
' Lines: 11
' ============================================================

Option Compare Database

Public Function CodeBehind_Form_Load() As Boolean
    lblMessage.Caption = message
    lblMessage2.Caption = Message1
    CodeBehind_Form_Load = True
End Function

Private Sub Form_Load()
CodeBehind_Form_Load
End Sub