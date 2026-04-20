# Report Lineage: New Invoice

## Trigger Paths
- frm_invoices_summary -> Cmd_PreviewNew -> onClick -> Cmd_PreviewNew_Click (high confidence)
- frmClientLedger -> cmdInvoice -> onClick -> cmdInvoice_Click (high confidence)
- zClient Ledger OLD -> cmdInvoice -> onClick -> cmdInvoice_Click (high confidence)

## Data Lineage
- RecordSource: `qry_current_invoice`
- RecordSourceType: `saved-query`
- Involved Queries: qry_current_invoice
- Terminal Tables: vw_current_invoice

## Related VBA Procedures
- form frm_invoices_summary::Cmd_PreviewNew_Click [8176-8206]
- form frm_invoices_summary::Cmd_PrintNew_Click [8207-8238]
- form frmClientLedger::cmdInvoice_Click [17642-17671]
- form zClient Ledger OLD::cmdInvoice_Click [7925-7945]
- report rpt_Compr_InvoiceADVCur::Report_Open [1610-1646]
