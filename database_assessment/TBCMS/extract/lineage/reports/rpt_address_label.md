# Report Lineage: rpt_address_label

## Trigger Paths
- frmClientLedger -> cmdPrintLabelP -> onClick -> cmdPrintLabelP_Click (high confidence)
- frmClientLedger -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frm_trust_summary -> cmbAddressLabel -> onClick -> cmbAddressLabel_Click (high confidence)
- frm_invoices_summary -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frmTimeKeepingClosed -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- zClient Ledger OLD -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16616-16646] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16647-16677] [createObject, runSql]
- form frmClientLedger::cmdPrintLabelP_Click [16884-16914]
- form frmClientLedger::cmdPrintLabel_Click [17484-17536] [runSql, setWarnings]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17547-17684] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17729-17739]
- form frm_trust_summary::cmbAddressLabel_Click [6871-6900]
- form frm_invoices_summary::cmdPrintLabel_Click [8301-8329]
- form frmTimeKeepingClosed::cmdPrintLabel_Click [7833-7862]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdPrintLabel_Click [8071-8086]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module DocumentManagement::MoveDocumentByCaseStatus [1373-1585] [createObject, fileSystem]
- module DocumentManagement::GetIntakeDocumentFileName [1682-1758]
- module DocumentManagement::Phase5_E2E_HappyPathTest [1759-2092] [createObject, fileSystem]
- module Module1::GetRetainer [20-23]
