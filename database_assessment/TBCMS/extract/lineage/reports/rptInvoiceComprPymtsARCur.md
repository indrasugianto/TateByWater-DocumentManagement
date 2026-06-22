# Report Lineage: rptInvoiceComprPymtsARCur

## Trigger Paths
- frmTimeKeepingClosed -> cmdCompCurr -> onClick -> cmdCompCurr_Click (medium confidence)
- Time Keeping -> cmdCompCurrent -> onClick -> cmdCompCurrent_Click (medium confidence)

## Data Lineage
- RecordSource: `qry_InvoicePymts_curr`
- RecordSourceType: `saved-query`
- Involved Queries: qry_InvoicePymts_curr, qry_current_invoice
- Terminal Tables: vw_current_invoice

## Related VBA Procedures
- form frmClientLedger::cmdInvoice_Click [17843-17872]
- form frm_invoices_summary::Cmd_PreviewNew_Click [8108-8138]
- form frmTimeKeepingClosed::cmdCompCurr_Click [7680-7728]
- form Time Keeping::cmdCompCurrent_Click [4846-4899]
- report rpt_Compr_InvoiceADVCur::Report_Open [1528-1564]
- report rptInvoiceComprPymtsARCur::Report_Open [515-519]
