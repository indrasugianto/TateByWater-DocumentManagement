# Report Lineage: rptInvoiceComprARCur

## Trigger Paths
- frmTimeKeepingClosed -> cmdCompCurr -> onClick -> cmdCompCurr_Click (high confidence)
- Time Keeping -> cmdCompCurrent -> onClick -> cmdCompCurrent_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT * FROM qry_InvoiceAR_curr;`
- RecordSourceType: `inline-sql`
- Involved Queries: qry_InvoiceAR_curr, qry_current_invoice
- Terminal Tables: vw_current_invoice

## Related VBA Procedures
- form frm_invoices_summary::Cmd_PreviewNew_Click [8176-8206]
- form frmClientLedger::cmdInvoice_Click [17642-17671]
- form frmTimeKeepingClosed::cmdCompCurr_Click [7742-7790]
- form Time Keeping::cmdCompCurrent_Click [4787-4840]
- report rptInvoiceComprARCur::Report_Open [584-590]
- report rpt_Compr_InvoiceADVCur::Report_Open [1610-1646]
