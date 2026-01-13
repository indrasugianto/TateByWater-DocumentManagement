' Component: Form_frmBilling
' Type: document
' Lines: 5
' ============================================================

Option Compare Database

Private Sub Past_Due_AfterUpdate()
    If Me.Dirty Then Me.Dirty = False
End Sub