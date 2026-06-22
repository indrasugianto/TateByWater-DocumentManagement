# Report Lineage: rptPIStatusSOL

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `qryAttyTrustAcctsTOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryAttyTrustAcctsTOff
- Terminal Tables: tblCase, tblTakeOff

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmTakeOffReconciliation::cmdInsertData_Click [3471-3550]
- form frmAttyFeeGeneration::FilterMe [8493-8505]
- form frmTakeOff::cmdInsertIntoTA_Click [4379-4453]
- form frmTakeOff::cmbInsertFees_Click [4890-4945]
- form frmTakeOff::fncRecordExists [4946-5010]
- form frmTakeOff2::cmdInsertIntoTA_Click [692-766]
- form frmTakeOff2::cmdShowReportTDT_Click [1140-1214]
- form frmTakeOff2::fncRecordExists [1215-1279]
- form frmTakeOffTest::cmdInsertIntoTA_Click [692-766]
- form frmTakeOffTest::cmbInsertFees_Click [1161-1214]
- form frmTakeOffTest::fncRecordExists [1215-1279]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
