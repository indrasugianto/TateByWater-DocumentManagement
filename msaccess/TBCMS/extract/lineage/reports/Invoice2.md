# Report Lineage: Invoice2

## Trigger Paths
- frm_invoices_summary -> cmdPreview -> onClick -> cmdPreview_Click (high confidence)
- frm_invoices_summary -> cmdRecordInvoice -> onClick -> cmdRecordInvoice_Click (high confidence)
- frmClientLedger -> CommandFullHistoryInvoice -> onClick -> CommandFullHistoryInvoice_Click (high confidence)
- frmClientLedger -> cmdEmailFullHistory -> onClick -> cmdEmailFullHistory_Click (high confidence)
- frmInvoiceSent -> cmdRecordSentInvoice -> onClick -> cmdRecordSentInvoice_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceRPT1
- Terminal Tables: vwInvoiceRPT1

## Related VBA Procedures
- form frm_invoices_summary::cmdPreview_Click [8308-8319]
- form frm_invoices_summary::cmdRecordInvoice_Click [8413-8471]
- form frmClientLedger::cmdEmailFullHistory_Click [16965-16988]
- form frmClientLedger::CommandFullHistoryInvoice_Click [17622-17641]
- form frmInvoiceSent::cmdRecordSentInvoice_Click [1115-1179]
- report Invoice2::Charge19_Click [2001-2033]
- report Rpt_MergeInvTK::Charge19_Click [1949-1981]
