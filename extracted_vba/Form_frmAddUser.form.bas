' Component: Form_frmAddUser
' Type: document
' Lines: 21
' ============================================================

Option Compare Database

Private Sub cmdNo_Click()

    ShowYesNo "The record has not been saved. Do you want to discard the changes and exit?"

    If YesNo_Value = 1 Then
        Me.Undo
        Util.CloseForm
    End If
End Sub

Private Sub cmdYes_Click()
    Util.CloseForm
End Sub

Public Function CodeBehind_Form_Load() As Boolean
    CodeBehind_Form_Load = True
End Function

