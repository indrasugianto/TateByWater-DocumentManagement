# Report Lineage: rpt_address_labelEx

## Trigger Paths
- frmClientLedger -> cmdPrintLabelP -> onClick -> cmdPrintLabelP_Click (high confidence)
- frmClientLedger -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frm_trust_summary -> cmbAddressLabel -> onClick -> cmbAddressLabel_Click (high confidence)
- frm_invoices_summary -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frmTimeKeepingClosed -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmdPrintLabelP_Click [16896-16926]
- form frmClientLedger::cmdPrintLabel_Click [17496-17548]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frm_trust_summary::cmbAddressLabel_Click [6871-6900]
- form frm_invoices_summary::cmdPrintLabel_Click [8301-8329]
- form frmTimeKeepingClosed::cmdPrintLabel_Click [7833-7862]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
