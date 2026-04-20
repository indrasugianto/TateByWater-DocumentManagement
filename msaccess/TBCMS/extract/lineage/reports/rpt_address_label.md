# Report Lineage: rpt_address_label

## Trigger Paths
- frm_invoices_summary -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frm_trust_summary -> cmbAddressLabel -> onClick -> cmbAddressLabel_Click (high confidence)
- frmClientLedger -> cmdPrintLabelP -> onClick -> cmdPrintLabelP_Click (high confidence)
- frmClientLedger -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- zClient Ledger OLD -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frmTimeKeepingClosed -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frm_invoices_summary::cmdPrintLabel_Click [8369-8397]
- form frm_trust_summary::cmbAddressLabel_Click [6891-6920]
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmdPrintLabelP_Click [16715-16745]
- form frmClientLedger::cmdPrintLabel_Click [17295-17347]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdPrintLabel_Click [8050-8065]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmTimeKeepingClosed::cmdPrintLabel_Click [7895-7924]
