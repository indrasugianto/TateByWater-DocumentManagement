# Report Lineage: rptReconciliation

## Trigger Paths
- frmTakeOff -> cmd_PreviewReconReport -> onClick -> cmd_PreviewReconReport_Click (high confidence)

## Data Lineage
- RecordSource: `tblTakeOffMonth`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblTakeOffMonth

## Related VBA Procedures
- form frmTakeOffReconciliation::cmdInsertData_Click [3495-3574] [runSql]
- form frmAttyFeeGeneration::FilterMe [8493-8505]
- form frmTakeOff::cmd_PreviewReconReport_Click [4472-4481]
- form frmTakeOff::cmbInsertFees_Click [4890-4945]
- form frmTakeOff::fncRecordExists [4946-5010] [runSql]
- form frmTakeOff2::cmd_PreviewReconReport_Click [785-794]
- form frmTakeOff2::cmdShowReportTDT_Click [1140-1214]
- form frmTakeOff2::fncRecordExists [1215-1279] [runSql]
- form frmTakeOffTest::cmd_PreviewReconReport_Click [785-794]
- form frmTakeOffTest::cmbInsertFees_Click [1161-1214]
- form frmTakeOffTest::fncRecordExists [1215-1279] [runSql]
