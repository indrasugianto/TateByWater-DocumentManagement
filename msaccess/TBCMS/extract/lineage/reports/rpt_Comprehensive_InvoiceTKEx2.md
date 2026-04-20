# Report Lineage: rpt_Comprehensive_InvoiceTKEx2

## Trigger Paths
- frmTimeKeepingClosed -> cmdRecord -> onClick -> cmdRecord_Click (high confidence)
- frmTimeKeepingClosed -> cmdRecordShort -> onClick -> cmdRecordShort_Click (medium confidence)
- Time Keeping -> cmdRecordTKStatement -> onClick -> cmdRecordTKStatement_Click (high confidence)
- Time Keeping -> cmdCompFullHistory -> onClick -> cmdCompFullHistory_Click (high confidence)
- Time Keeping -> cmdCompShort -> onClick -> cmdCompShort_Click (medium confidence)
- Time Keeping -> cmdEmailShort -> onClick -> cmdEmailShort_Click (medium confidence)
- Time Keeping -> cmdEmailLong -> onClick -> cmdEmailLong_Click (high confidence)
- Time Keeping -> cmdRecordShortTK -> onClick -> cmdRecordShortTK_Click (medium confidence)

## Data Lineage
- RecordSource: `SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: TB Time Keeping, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form frmTakeOffReconciliation::txtTKButton_Click [3724-3781]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7452-7533]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8008-8032]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7925-8041]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8062-8104]
- form frmTimeKeepingClosed::cmdRecord_Click [8105-8221]
- form Time Keeping::cmdCompFullHistory_Click [3989-4030]
- form Time Keeping::cmdCompShort_Click [4031-4073]
- form Time Keeping::cmdEmailLong_Click [4074-4119]
- form Time Keeping::cmdEmailShort_Click [4120-4165]
- form Time Keeping::cmdRecordShortTK_Click [4199-4316]
- form Time Keeping::cmdRecordTKStatement_Click [4457-4572]
- form Time Keeping::cmdCreateAR_Click [4573-4608]
- form Time Keeping::cmdInsertTime_Click [4628-4651]
- form Time Keeping::cmdAddNew_Click [4652-4673]
- form Time Keeping::addNewTK [4674-4729]
- form frmTKClose::txtTKButton_Click [2512-2841]
