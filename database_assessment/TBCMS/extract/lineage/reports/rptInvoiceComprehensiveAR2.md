# Report Lineage: rptInvoiceComprehensiveAR2

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, tblCase.Retainer, tblCase.CaseOpenDate, [Matter and AR].OrderNr, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Bill_ID FROM (tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID) INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID WHERE ((([Matter and AR].Date2)<=[Bill Closed Date]) AND (([Matter and AR].Charge)<>0)) ORDER BY [Matter and AR].Date2;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Matter and AR, TB Time Keeping

## Related VBA Procedures
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frm_advanced_payments::FilterMe [6919-6945]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7452-7533] [runSql]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389] [runSql]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8008-8032]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3103-3203] [runSql]
- form frmMatter::Date2_AfterUpdate [1513-1563] [runSql]
- form frmMatter::reorderByDateMatter [1572-1593] [runSql]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7925-8041] [outputTo, runSql]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8062-8104] [runSql]
- form frmTimeKeepingClosed::cmdRecord_Click [8105-8221] [outputTo, runSql]
- form Time Keeping::cmdRecordShortTK_Click [4238-4355] [outputTo, runSql]
- form Time Keeping::cmdRecordTKStatement_Click [4517-4632] [outputTo, runSql]
- form Time Keeping::cmdCreateAR_Click [4633-4668] [runSql]
- form Time Keeping::cmdInsertTime_Click [4688-4711] [runSql]
- form Time Keeping::cmdAddNew_Click [4712-4733]
- form Time Keeping::addNewTK [4734-4789] [runSql]
- form frmTKClose::txtTKButton_Click [2378-2707] [runSql]
- form frmTakeOffSubForm3::CaseNum_Click [2558-2680] [runSql]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2653-2762] [runSql]
- report Invoice2::Charge19_Click [2068-2100] [runSql]
- report Rpt_MergeInvTK::Charge19_Click [1949-1981] [runSql]
- module modGaz::fncRunningDebit [2-17]
- module modGaz::fncRunningCredit [18-33]
- module modGaz::fnc_TEST_get_remaining_AdvancedChargesBalance [399-405]
