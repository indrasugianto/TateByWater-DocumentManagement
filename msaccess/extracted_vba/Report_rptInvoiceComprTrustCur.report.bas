' Component: Report_rptInvoiceComprTrustCur
' Type: document
' Lines: 29
' ============================================================

Option Compare Database

Private Sub Report_NoData(Cancel As Integer)
    Cancel = True
End Sub

'Private Sub Report_Open(Cancel As Integer)
'    intCaseID = Nz(Form_frmClientLedger.CaseID, 0)
'    strFilter = "CaseID=" & intCaseID & " AND Balance=0"
'    intOrderNr = DMax("OrderNr", "qry_invoice_comprehensive_trust_acc_cur", strFilter)
'
'    If IsNull(intOrderNr) Then
'        strOrderNr = ""
'    Else
'        strOrderNr = " AND OrderNr>" & intOrderNr
'    End If
'
'    strFilter = "[CaseID]=" & Nz(Form_frmClientLedger.CaseID, 0) & strOrderNr
'    On Error Resume Next
'    Me.filter = strFilter
'
'    Me.OrderBy = "OrderNr"
'End Sub

Private Sub Report_Open(Cancel As Integer)
    strSQLRecSource = "select * from qryInvoiceComprehensiveTrustCredit where CaseID=" & Form_frmClientLedger.CaseID & " AND " & fncGetFilterOrderNrTA(Form_frmClientLedger.CaseID)
    Debug.Print strSQLRecSource
    Me.RecordSource = strSQLRecSource
End Sub