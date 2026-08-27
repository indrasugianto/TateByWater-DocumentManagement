# Report Lineage: rpt_address_label

## Trigger Paths
- frmClientLedger -> cmdPrintLabelP -> onClick -> cmdPrintLabelP_Click (high confidence)
- frmClientLedger -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frm_trust_summary -> cmbAddressLabel -> onClick -> cmbAddressLabel_Click (high confidence)
- frm_invoices_summary -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- zClient Ledger OLD -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)
- frmTimeKeepingClosed -> cmdPrintLabel -> onClick -> cmdPrintLabel_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmdPrintLabelP_Click [18700-18730]
- form frmClientLedger::cmdPrintLabel_Click [19300-19352] [runSql, setWarnings]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form frm_trust_summary::cmbAddressLabel_Click [6871-6900]
- form frm_invoices_summary::cmdPrintLabel_Click [8301-8329]
- form zClient Ledger OLD::cmdPrintLabel_Click [8050-8065]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmTimeKeepingClosed::cmdPrintLabel_Click [7895-7924]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
