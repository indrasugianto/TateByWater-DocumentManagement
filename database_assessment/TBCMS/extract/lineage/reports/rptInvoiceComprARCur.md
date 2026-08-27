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
- form frmClientLedger::cmdInvoice_Click [19647-19676]
- form frm_invoices_summary::Cmd_PreviewNew_Click [8108-8138]
- form frmTimeKeepingClosed::cmdCompCurr_Click [7742-7790]
- form Time Keeping::cmdCompCurrent_Click [4847-4900]
- report rptInvoiceComprARCur::Report_Open [584-590]
- report rpt_Compr_InvoiceADVCur::Report_Open [1610-1646]
- module modGaz::fncGetFilterOrderNrMatterAR [194-207]
