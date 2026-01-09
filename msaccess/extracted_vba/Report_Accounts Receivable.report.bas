' Component: Report_Accounts Receivable
' Type: document
' Lines: 17
' ============================================================

Option Compare Database

'Private Sub Report_Open(Cancel As Integer)
'    'Debug.Print DLookup("SumOfRetBal", "qry_RetBalSums_by_PastDue", "Orig_Atty='" & Form_frm_invoices_summary.CmbOrigAtty & "' and [Past Due] = 0")
'    NonPastDueValue = DLookup("SumOfRetBal", "qry_RetBalSums_by_PastDue", "Orig_Atty='" & Form_frm_invoices_summary.cmbOrigAtty & "' and [Past Due] = 0")
'    NonPastDueValue = Nz(NonPastDueValue, 0)
'    Me.txtNonPastDue.ControlSource = "=" & NonPastDueValue
'
'
'    PastDueValue = DLookup("SumOfRetBal", "qry_RetBalSums_by_PastDue", "Orig_Atty='" & Form_frm_invoices_summary.cmbOrigAtty & "' and [Past Due] = -1")
'    PastDueValue = Nz(PastDueValue, 0)
'    Me.txtPastDue.ControlSource = "=" & PastDueValue
'End Sub

Private Sub Report_Load()
    'DoCmd.ApplyFilter , "BalRetCalculated <> 0"
End Sub