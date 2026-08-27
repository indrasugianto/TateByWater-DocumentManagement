# Report Lineage: Invoice - No Balance Due

## Trigger Paths
- frmClientLedger -> cmdBalanceInvoice -> onClick -> cmdBalanceInvoice_Click (high confidence)
- frmClientLedger -> cmdEmailNoBalance -> onClick -> cmdEmailNoBalance_Click (high confidence)
- zClient Ledger OLD -> cmdBalanceInvoice -> onClick -> cmdBalanceInvoice_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceRPT1
- Terminal Tables: vwInvoiceRPT1

## Related VBA Procedures
- form frmClientLedger::cmdEmailNoBalance_Click [18946-18969]
- form frmClientLedger::cmdBalanceInvoice_Click [19243-19299]
- form frm_invoices_summary::cmdNoBalancePrint_Click [8281-8296]
- form zClient Ledger OLD::cmdBalanceInvoice_Click [7965-7981]
- form frmHome::cmdDeleteAll_Click [1868-2005] [runSql]
- form frmHomeAdmin::cmdDeleteAll_Click [6313-6450] [runSql]
