# Report Lineage: rptPIStatusSOL

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `qryAttyTrustAcctsTOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryAttyTrustAcctsTOff
- Terminal Tables: tblCase, tblTakeOff

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form frmAttyFeeGeneration::FilterMe [8563-8575]
- form frmTakeOffReconciliation::cmdInsertData_Click [3471-3550] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmTakeOff::cmdInsertIntoTA_Click [4378-4452] [runSql]
- form frmTakeOff::cmbInsertFees_Click [4889-4944]
- form frmTakeOff::fncRecordExists [4945-5009] [runSql]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmTakeOff2::cmdInsertIntoTA_Click [674-748] [runSql]
- form frmTakeOff2::cmdShowReportTDT_Click [1122-1196]
- form frmTakeOff2::fncRecordExists [1197-1261] [runSql]
- form frmTakeOffTest::cmdInsertIntoTA_Click [675-749] [runSql]
- form frmTakeOffTest::cmbInsertFees_Click [1144-1197]
- form frmTakeOffTest::fncRecordExists [1198-1262] [runSql]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
