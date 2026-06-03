# Report Lineage: rpt_Comprehensive_InvoiceStmtS

## Trigger Paths
- frmTimeKeepingClosed -> cmdRecordShort -> onClick -> cmdRecordShort_Click (high confidence)
- Time Keeping -> cmdCompShort -> onClick -> cmdCompShort_Click (high confidence)
- Time Keeping -> cmdEmailShort -> onClick -> cmdEmailShort_Click (high confidence)
- Time Keeping -> cmdRecordShortTK -> onClick -> cmdRecordShortTK_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, [TB Time Keeping].AdvBalanceatClose, [TB Time Keeping].TrustatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: TB Time Keeping, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389] [runSql]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7863-7979] [outputTo, runSql]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8000-8042] [runSql]
- form frmTimeKeepingClosed::cmdRecord_Click [8043-8159] [outputTo, runSql]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7416-7497] [runSql]
- form frmTKClose::txtTKButton_Click [2379-2708] [runSql]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form Time Keeping::cmdCompShort_Click [4057-4099]
- form Time Keeping::cmdEmailShort_Click [4146-4203]
- form Time Keeping::cmdRecordShortTK_Click [4237-4354] [outputTo, runSql]
- form Time Keeping::cmdRecordTKStatement_Click [4516-4631] [outputTo, runSql]
- form Time Keeping::cmdCreateAR_Click [4632-4667] [runSql]
- form Time Keeping::cmdInsertTime_Click [4687-4710] [runSql]
- form Time Keeping::cmdAddNew_Click [4711-4732]
- form Time Keeping::addNewTK [4733-4788] [runSql]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8029-8053]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
