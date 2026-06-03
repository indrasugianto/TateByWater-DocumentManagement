# Report Lineage: rpt_opp_counsel_address_label

## Trigger Paths
- frmClientLedger -> cmdPrintOCLabel -> onClick -> cmdPrintOCLabel_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689] [createObject, runSql]
- form frmClientLedger::cmdPrintOCLabel_Click [16927-16942]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
