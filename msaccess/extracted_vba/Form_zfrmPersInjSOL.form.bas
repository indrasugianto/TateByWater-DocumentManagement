' Component: Form_zfrmPersInjSOL
' Type: document
' Lines: 5
' ============================================================

Option Compare Database

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub