' Component: Report_rptInvoiceComprARCur
' Type: document
' Lines: 25
' ============================================================

Option Compare Database

'Private Sub Report_Open(Cancel As Integer)
'    intCaseID = Nz(Form_frmClientLedger.CaseID, 0)
'    strFilter = "CaseID=" & intCaseID & " AND Balance=0"
'    intOrderNr = DMax("OrderNr", "qry_current_invoice", strFilter)
'
'    If IsNull(intOrderNr) Then
'        strOrderNr = ""
'    Else
'        strOrderNr = " AND OrderNr>" & intOrderNr
'    End If
'
'    strFilter = "[CaseID]=" & Nz(Form_frmClientLedger.CaseID, 0) & strOrderNr
'    Me.filter = strFilter
'
'    Me.OrderBy = "OrderNr"
'End Sub
Private Sub Report_Open(Cancel As Integer)
    strSQLRecSource = "select * from qry_InvoiceAR_curr where CaseID=" & Form_frmClientLedger.CaseID & " AND " & fncGetFilterOrderNrMatterAR(Form_frmClientLedger.CaseID)
    Debug.Print strSQLRecSource
    Me.RecordSource = strSQLRecSource
    'Me.Requery
    'Beep
End Sub