# Report Lineage: Invoice - Past Due

## Trigger Paths
- frmClientLedger -> cmdPastDueInvoice -> onClick -> cmdPastDueInvoice_Click (high confidence)
- frmClientLedger -> cmdEmailPastDue -> onClick -> cmdEmailPastDue_Click (high confidence)
- frm_invoices_summary -> cmdPastDue -> onClick -> cmdPastDue_Click (high confidence)
- frm_invoices_summary -> cmdPastDueLT -> onClick -> cmdPastDueLT_Click (high confidence)
- frm_invoices_summary -> cmdRecordPDInvoice -> onClick -> cmdRecordPDInvoice_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmInvoiceSent -> cmdRecondSentPDInvoice -> onClick -> cmdRecondSentPDInvoice_Click (high confidence)
- zClient Ledger OLD -> cmdPastDueInvoice -> onClick -> cmdPastDueInvoice_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceRPT1
- Terminal Tables: vwInvoiceRPT1

## Related VBA Procedures
- form frmClientLedger::cmdEmailPastDue_Click [17115-17141]
- form frmClientLedger::cmdPastDueInvoice_Click [17411-17438]
- form frm_invoices_summary::cmdPastDueLT_Click [8191-8207]
- form frm_invoices_summary::cmdPastDue_Click [8224-8239]
- form frm_invoices_summary::cmdPastDuePrint_Click [8266-8280]
- form frm_invoices_summary::cmdPrintPastDueLT_Click [8330-8344]
- form frm_invoices_summary::cmdRecordPDInvoice_Click [8472-8530] [outputTo, runSql]
- form frmHomeAdmin::cmdDeleteAll_Click [6313-6450] [runSql]
- form frmHome::cmdDeleteAll_Click [1868-2005] [runSql]
- form frmInvoiceSent::cmdRecondSentPDInvoice_Click [1015-1077] [outputTo, runSql]
- form zClient Ledger OLD::cmdPastDueInvoice_Click [7967-7985]
