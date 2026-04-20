# Report Lineage: rptInvoiceComprehensiveAR2

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, tblCase.Retainer, tblCase.CaseOpenDate, [Matter and AR].OrderNr, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Bill_ID FROM (tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID) INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID WHERE ((([Matter and AR].Date2)<=[Bill Closed Date]) AND (([Matter and AR].Charge)<>0)) ORDER BY [Matter and AR].Date2;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Matter and AR, TB Time Keeping

## Related VBA Procedures
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frm_advanced_payments::FilterMe [7008-7035]
- form frmTakeOffReconciliation::txtTKButton_Click [3724-3781]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7452-7533]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8008-8032]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3192-3292]
- form frmMatter::Date2_AfterUpdate [1513-1563]
- form frmMatter::reorderByDateMatter [1572-1593]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7925-8041]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8062-8104]
- form frmTimeKeepingClosed::cmdRecord_Click [8105-8221]
- form Time Keeping::cmdRecordShortTK_Click [4199-4316]
- form Time Keeping::cmdRecordTKStatement_Click [4457-4572]
- form Time Keeping::cmdCreateAR_Click [4573-4608]
- form Time Keeping::cmdInsertTime_Click [4628-4651]
- form Time Keeping::cmdAddNew_Click [4652-4673]
- form Time Keeping::addNewTK [4674-4729]
- form frmTKClose::txtTKButton_Click [2512-2841]
- form frmTakeOffSubForm3::CaseNum_Click [2558-2680]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2653-2762]
- report Invoice2::Charge19_Click [2001-2033]
- report Rpt_MergeInvTK::Charge19_Click [1949-1981]
