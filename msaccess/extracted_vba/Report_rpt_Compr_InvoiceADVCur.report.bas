' Component: Report_rpt_Compr_InvoiceADVCur
' Type: document
' Lines: 41
' ============================================================

Option Compare Database

'PARENT REPORT!!!!!!!!!!!!!!!!!!!!

Private Sub Report_Open(Cancel As Integer)
    
    'rptInvoiceComprPymtsARCur.Report.filter = fncGetFilterOrderNr(Me.CaseID)
    
    
    
'    strSQLRecSource = "select * from qry_InvoiceAR_curr where CaseID=" & Form_frmClientLedger.CaseID & " AND " & fncGetFilterOrderNr(Form_frmClientLedger.CaseID)
'    Debug.Print strSQLRecSource
'    rptInvoiceComprPymtsARCur.Report.RecordSource = strSQLRecSource
    
    
    
    
'    'rptInvoiceComprPymtsARCur BEGIN
'        intCaseID = Nz(Form_frmClientLedger.CaseID, 0)
'        strFilter = "CaseID=" & intCaseID & " AND Balance=0"
'        intOrderNr = DMax("OrderNr", "qry_current_invoice", strFilter)
'
'        If IsNull(intOrderNr) Then
'            strOrderNr = ""
'        Else
'            strOrderNr = "[OrderNr] > " & intOrderNr
'        End If
'
'        strFilter = strOrderNr
'        rptInvoiceComprPymtsARCur.Report.filter = strFilter
'        rptInvoiceComprPymtsARCur.Report.OrderBy = "OrderNr"
'    'rptInvoiceComprPymtsARCur END
'
End Sub

'Private Sub Report_Open(Cancel As Integer)
'    If Report_rptInvoiceComprehensiveTrust.Report.Recordset.EOF Then
'        Debug.Print "Empty"
'    End If
'End Sub
