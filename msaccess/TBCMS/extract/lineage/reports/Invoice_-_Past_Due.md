# Report Lineage: Invoice - Past Due

## Trigger Paths
- frm_invoices_summary -> cmdPastDue -> onClick -> cmdPastDue_Click (high confidence)
- frm_invoices_summary -> cmdPastDueLT -> onClick -> cmdPastDueLT_Click (high confidence)
- frm_invoices_summary -> cmdRecordPDInvoice -> onClick -> cmdRecordPDInvoice_Click (high confidence)
- frmClientLedger -> cmdPastDueInvoice -> onClick -> cmdPastDueInvoice_Click (high confidence)
- frmClientLedger -> cmdEmailPastDue -> onClick -> cmdEmailPastDue_Click (high confidence)
- zClient Ledger OLD -> cmdPastDueInvoice -> onClick -> cmdPastDueInvoice_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmInvoiceSent -> cmdRecondSentPDInvoice -> onClick -> cmdRecondSentPDInvoice_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceRPT1
- Terminal Tables: vwInvoiceRPT1

## Related VBA Procedures
- form frm_invoices_summary::cmdPastDueLT_Click [8259-8275]
- form frm_invoices_summary::cmdPastDue_Click [8292-8307]
- form frm_invoices_summary::cmdPastDuePrint_Click [8334-8348]
- form frm_invoices_summary::cmdPrintPastDueLT_Click [8398-8412]
- form frm_invoices_summary::cmdRecordPDInvoice_Click [8540-8598]
- form frmClientLedger::cmdEmailPastDue_Click [16914-16940]
- form frmClientLedger::cmdPastDueInvoice_Click [17210-17237]
- form zClient Ledger OLD::cmdPastDueInvoice_Click [7946-7964]
- form frmHome::cmdDeleteAll_Click [1920-2057]
- form frmInvoiceSent::cmdRecondSentPDInvoice_Click [1052-1114]
- form frmHomeAdmin::cmdDeleteAll_Click [6356-6493]
