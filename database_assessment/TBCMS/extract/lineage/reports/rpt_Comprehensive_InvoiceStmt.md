# Report Lineage: rpt_Comprehensive_InvoiceStmt

## Trigger Paths
- frmTimeKeepingClosed -> cmdRecord -> onClick -> cmdRecord_Click (high confidence)
- frmTimeKeepingClosed -> cmdCompFull -> onClick -> cmdCompFull_Click (high confidence)
- frmTimeKeepingClosed -> cmdRecordShort -> onClick -> cmdRecordShort_Click (medium confidence)
- Time Keeping -> cmdRecordTKStatement -> onClick -> cmdRecordTKStatement_Click (high confidence)
- Time Keeping -> cmdCompFullHistory -> onClick -> cmdCompFullHistory_Click (high confidence)
- Time Keeping -> cmdCompShort -> onClick -> cmdCompShort_Click (medium confidence)
- Time Keeping -> cmdEmailShort -> onClick -> cmdEmailShort_Click (medium confidence)
- Time Keeping -> cmdEmailLong -> onClick -> cmdEmailLong_Click (high confidence)
- Time Keeping -> cmdRecordShortTK -> onClick -> cmdRecordShortTK_Click (medium confidence)

## Data Lineage
- RecordSource: `SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, [TB Time Keeping].AdvBalanceatClose, [TB Time Keeping].TrustatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: TB Time Keeping, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7452-7533] [runSql]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389] [runSql]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8008-8032]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmTimeKeepingClosed::cmdCompFull_Click [7791-7823]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7925-8041] [outputTo, runSql]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8062-8104] [runSql]
- form frmTimeKeepingClosed::cmdRecord_Click [8105-8221] [outputTo, runSql]
- form Time Keeping::cmdCompFullHistory_Click [4016-4057]
- form Time Keeping::cmdCompShort_Click [4058-4100]
- form Time Keeping::cmdEmailLong_Click [4101-4146]
- form Time Keeping::cmdEmailShort_Click [4147-4204]
- form Time Keeping::cmdRecordShortTK_Click [4238-4355] [outputTo, runSql]
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
