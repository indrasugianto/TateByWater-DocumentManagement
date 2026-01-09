' Component: Form_frmTakeOffSteps
' Type: document
' Lines: 29
' ============================================================

Option Compare Database

Private Sub cmdAdvanced_Click()
    DoCmd.openform "frm_advanced_payments"
End Sub

Private Sub cmdAttyFeesCost_Click()
    DoCmd.openform "frmattyfeegeneration"
End Sub

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdRecon_Click()
    DoCmd.openform "frmtakeoffreconciliation"
End Sub

Private Sub cmdTakeOff_Click()
    DoCmd.openform "frmtakeoff"
End Sub

Private Sub cmdTrsutChron_Click()
    DoCmd.openform "frmtrustentrieschron"
End Sub

Private Sub cmdUncashed_Click()
    DoCmd.openform "frm_uncashed_trust_checks"
End Sub