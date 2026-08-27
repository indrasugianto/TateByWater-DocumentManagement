# Report Lineage: Accounts Receivable

## Trigger Paths
- frm_invoices_summary -> cmdOpenReportAcctReceivable -> onClick -> cmdOpenReportAcctReceivable_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)

## Data Lineage
- RecordSource: `qry_invoices_summaryRPT`
- RecordSourceType: `saved-query`
- Involved Queries: qry_invoices_summaryRPT, qryInvoiceRPT1
- Terminal Tables: tblDropD, vwInvoiceRPT1

## Related VBA Procedures
- form frm_invoices_summary::cmdOpenReportAcctReceivable_Click [8457-8467]
- form frmHome::cmdDeleteAll_Click [1868-2005] [runSql]
- form frmHomeAdmin::cmdDeleteAll_Click [6313-6450] [runSql]
