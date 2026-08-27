# Report Lineage: rptReconciliation

## Trigger Paths
- frmTakeOff -> cmd_PreviewReconReport -> onClick -> cmd_PreviewReconReport_Click (high confidence)

## Data Lineage
- RecordSource: `tblTakeOffMonth`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblTakeOffMonth

## Related VBA Procedures
- form frmAttyFeeGeneration::FilterMe [8563-8575]
- form frmTakeOffReconciliation::cmdInsertData_Click [3471-3550] [runSql]
- form frmTakeOff::cmd_PreviewReconReport_Click [4471-4480]
- form frmTakeOff::cmbInsertFees_Click [4889-4944]
- form frmTakeOff::fncRecordExists [4945-5009] [runSql]
- form frmTakeOff2::cmd_PreviewReconReport_Click [767-776]
- form frmTakeOff2::cmdShowReportTDT_Click [1122-1196]
- form frmTakeOff2::fncRecordExists [1197-1261] [runSql]
- form frmTakeOffTest::cmd_PreviewReconReport_Click [768-777]
- form frmTakeOffTest::cmbInsertFees_Click [1144-1197]
- form frmTakeOffTest::fncRecordExists [1198-1262] [runSql]
