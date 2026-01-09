' Component: Form_frmUsers
' Type: document
' Lines: 12
' ============================================================

Option Compare Database

Private Sub Form_Load()

    Dim passed As Boolean

    passed = FormUtils.Form_Load_New(Me, True)
    
    On Error Resume Next
    If passed = False Then DoCmd.Close acForm, Me.Name

End Sub