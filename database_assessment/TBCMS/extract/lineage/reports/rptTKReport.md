# Report Lineage: rptTKReport

## Trigger Paths
- frmClientLedger -> cmdTKReportAR -> onClick -> cmdTKReportAR_Click (medium confidence)
- Time Keeping -> cmdTKReport -> onClick -> cmdTKReport_Click (medium confidence)

## Data Lineage
- RecordSource: `qryTimeKeeping`
- RecordSourceType: `saved-query`
- Involved Queries: qryTimeKeeping
- Terminal Tables: TB Time Keeping, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmdTKReportAR_Click [18822-18841]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7452-7533] [runSql]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389] [runSql]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8008-8032]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7925-8041] [outputTo, runSql]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8062-8104] [runSql]
- form frmTimeKeepingClosed::cmdRecord_Click [8105-8221] [outputTo, runSql]
- form Time Keeping::cmdRecordShortTK_Click [4238-4355] [outputTo, runSql]
- form Time Keeping::cmdTKReport_Click [4365-4384]
- form Time Keeping::cmdRecordTKStatement_Click [4517-4632] [outputTo, runSql]
- form Time Keeping::cmdCreateAR_Click [4633-4668] [runSql]
- form Time Keeping::cmdInsertTime_Click [4688-4711] [runSql]
- form Time Keeping::cmdAddNew_Click [4712-4733]
- form Time Keeping::addNewTK [4734-4789] [runSql]
- form frmTKClose::txtTKButton_Click [2378-2707] [runSql]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
