' Component: Form_frmConflictChk
' Type: document
' Lines: 10
' ============================================================

Option Compare Database


Private Sub cmbclose_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmbHome_Click()
DoCmd.openform "frmhome", acNormal
End Sub