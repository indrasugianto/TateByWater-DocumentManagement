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
- form frmClientLedger::cmdInvoice_Click [17843-17872]
- form frm_invoices_summary::Cmd_PreviewNew_Click [8108-8138]
- form frmTimeKeepingClosed::cmdCompCurr_Click [7680-7728]
- form Time Keeping::cmdCompCurrent_Click [4846-4899]
- report rpt_Compr_InvoiceADVCur::Report_Open [1528-1564]
- report rptInvoiceComprARCur::Report_Open [525-531]
- module modGaz::fncGetFilterOrderNrMatterAR [194-207]
