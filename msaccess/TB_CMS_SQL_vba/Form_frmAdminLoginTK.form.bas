' Component: Form_frmAdminLoginTK
' Type: document
' Lines: 30
' ============================================================

Option Compare Database

Private Sub cmdNo_Click()
     DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdYes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdYes_Click
End Sub

Private Sub cmdYes_Click()

On Error GoTo errorhandler
If Me.txtpwd = "27admin40" Then
DoCmd.openform "frmTKClose", , , stlinkcriteria
DoCmd.Close acForm, Me.Name
Forms("frmTKClose").SetFocus
Else
DoCmd.Close acForm, Me.Name
MsgBox "Sorry, the password is incorrect.", , "TB CMS"
End If
Exit Sub
errorhandler:
MsgBox Err.Description, vbCritical, "Error #" & Err.Number

End Sub



