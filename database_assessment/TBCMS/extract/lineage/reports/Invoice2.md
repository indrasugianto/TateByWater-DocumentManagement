# Report Lineage: Invoice2

## Trigger Paths
- frmClientLedger -> CommandFullHistoryInvoice -> onClick -> CommandFullHistoryInvoice_Click (high confidence)
- frmClientLedger -> cmdEmailFullHistory -> onClick -> cmdEmailFullHistory_Click (high confidence)
- frm_invoices_summary -> cmdPreview -> onClick -> cmdPreview_Click (high confidence)
- frm_invoices_summary -> cmdRecordInvoice -> onClick -> cmdRecordInvoice_Click (high confidence)
- frmInvoiceSent -> cmdRecordSentInvoice -> onClick -> cmdRecordSentInvoice_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceRPT1
- Terminal Tables: vwInvoiceRPT1

## Related VBA Procedures
- form frmClientLedger::cmdEmailFullHistory_Click [18970-18993]
- form frmClientLedger::CommandFullHistoryInvoice_Click [19627-19646]
- form frm_invoices_summary::cmdPreview_Click [8240-8251]
- form frm_invoices_summary::cmdRecordInvoice_Click [8345-8403] [outputTo, runSql]
- form frmInvoiceSent::cmdRecordSentInvoice_Click [1075-1139] [outputTo, runSql]
- report Invoice2::Charge19_Click [2068-2100] [runSql]
- report Rpt_MergeInvTK::Charge19_Click [1949-1981] [runSql]
