# Report Lineage: rptPIStatusSOL

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `qryAttyTrustAcctsTOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryAttyTrustAcctsTOff
- Terminal Tables: tblCase, tblTakeOff

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form frmAttyFeeGeneration::FilterMe [8563-8575]
- form frmTakeOffReconciliation::cmdInsertData_Click [3863-3942]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmTakeOff::cmdInsertIntoTA_Click [5340-5414]
- form frmTakeOff::cmbInsertFees_Click [5851-5907]
- form frmTakeOff::fncRecordExists [5908-5972]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmTakeOff2::cmdInsertIntoTA_Click [674-748]
- form frmTakeOff2::cmdShowReportTDT_Click [1122-1196]
- form frmTakeOff2::fncRecordExists [1197-1261]
- form frmTakeOffTest::cmdInsertIntoTA_Click [675-749]
- form frmTakeOffTest::cmbInsertFees_Click [1144-1197]
- form frmTakeOffTest::fncRecordExists [1198-1262]
