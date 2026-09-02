# Report Lineage: rpt_Billing_Closing

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, Disposition.[Total Earned Fee] AS Expr1 FROM tblCase INNER JOIN Disposition ON tblCase.CaseID = Disposition.CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Disposition, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16616-16646] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16647-16677] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17547-17684] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17729-17739]
- form frmDispositions::FilterMe [7360-7390]
- form frmHomeAdmin::cmdDispositions_Click [6225-6228]
- form frmHomeAdmin::Command119_Click [6469-6472]
- form frmHome::cmdDispositions_Click [1764-1767]
- form frmHome::Command119_Click [1988-1991]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module DocumentManagement::MoveDocumentByCaseStatus [1373-1585] [createObject, fileSystem]
- module DocumentManagement::GetIntakeDocumentFileName [1682-1758]
- module DocumentManagement::Phase5_E2E_HappyPathTest [1759-2092] [createObject, fileSystem]
- module Module1::GetRetainer [20-23]
