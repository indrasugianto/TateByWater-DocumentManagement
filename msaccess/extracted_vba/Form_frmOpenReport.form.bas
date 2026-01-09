' Component: Form_frmOpenReport
' Type: document
' Lines: 13
' ============================================================

Option Compare Database

Private Sub cmdAllAttyReport_Click()
    DoCmd.Close acReport, "rptPersInjuryStatus", acSaveYes
    DoCmd.OpenReport "rptPersInjuryStatus", acViewPreview
    DoCmd.Close acForm, "frmPersInjuryStatusReport"
End Sub

Private Sub cmdAttyReport_Click()
    DoCmd.Close acReport, "rptPersInjuryStatus", acSaveYes
    DoCmd.OpenReport "rptPersInjuryStatus", acViewPreview, "", "Orig_Atty='" & Me.cmbPIAtty & "'"
    DoCmd.Close acForm, "frmPersInjuryStatusReport"
End Sub