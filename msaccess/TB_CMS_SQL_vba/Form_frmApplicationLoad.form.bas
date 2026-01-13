' Component: Form_frmApplicationLoad
' Type: document
' Lines: 8
' ============================================================

Option Compare Database

Private Sub Form_Load()
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.openform "frmLogin", acNormal
End Sub