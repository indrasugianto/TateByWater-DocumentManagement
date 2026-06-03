# Report Lineage: rpt_Disposition_Closing

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `qry_Disposition_ClosingSheet`
- RecordSourceType: `saved-query`
- Involved Queries: qry_Disposition_ClosingSheet
- Terminal Tables: Disposition, TblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmDispositions::FilterMe [7360-7390]
- form frmHomeAdmin::cmdDispositions_Click [6225-6228]
- form frmHomeAdmin::Command119_Click [6469-6472]
- form frmHome::cmdDispositions_Click [1800-1803]
- form frmHome::Command119_Click [2024-2027]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
