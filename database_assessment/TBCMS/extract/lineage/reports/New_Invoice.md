# Report Lineage: New Invoice

## Trigger Paths
- frmClientLedger -> cmdInvoice -> onClick -> cmdInvoice_Click (high confidence)
- frm_invoices_summary -> Cmd_PreviewNew -> onClick -> Cmd_PreviewNew_Click (high confidence)
- zClient Ledger OLD -> cmdInvoice -> onClick -> cmdInvoice_Click (high confidence)

## Data Lineage
- RecordSource: `qry_current_invoice`
- RecordSourceType: `saved-query`
- Involved Queries: qry_current_invoice
- Terminal Tables: vw_current_invoice

## Related VBA Procedures
- form frmClientLedger::cmdInvoice_Click [19647-19676]
- form frm_invoices_summary::Cmd_PreviewNew_Click [8108-8138]
- form frm_invoices_summary::Cmd_PrintNew_Click [8139-8170]
- form zClient Ledger OLD::cmdInvoice_Click [7925-7945]
- report rpt_Compr_InvoiceADVCur::Report_Open [1610-1646]
- module modGaz::fncGetFilterOrderNrMatterAR [194-207]
