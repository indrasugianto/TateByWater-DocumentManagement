# Report Lineage: rptInvoiceComprPymtsAR

## Trigger Paths
- frmTimeKeepingClosed -> cmdCompCurr -> onClick -> cmdCompCurr_Click (medium confidence)
- Time Keeping -> cmdCompCurrent -> onClick -> cmdCompCurrent_Click (medium confidence)

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Payment, tblCase.CaseOpenDate, [Matter and AR].OrderNr, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Bill_ID FROM (tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID) INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID WHERE ((([Matter and AR].Date2)<=[Bill Closed Date]) AND (([Matter and AR].Pay_Outlay) Not Like "Adjustment") AND (([Matter and AR].Payment)>0)) ORDER BY [Matter and AR].Date2;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Matter and AR, TB Time Keeping

## Related VBA Procedures
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17547-17684] [runSql]
- form frm_advanced_payments::FilterMe [6918-6944]
- form frmTakeOffReconciliation::txtTKButton_Click [3356-3413] [runSql]
- form frmMatter::Date2_AfterUpdate [1512-1562] [runSql]
- form frmMatter::reorderByDateMatter [1571-1592] [runSql]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3104-3204] [runSql]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2573-2682] [runSql]
- form frmTakeOffSubForm3::CaseNum_Click [2483-2605] [runSql]
- form frmTimeKeepingClosed::cmdCompCurr_Click [7680-7728]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7863-7979] [outputTo, runSql]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8000-8042] [runSql]
- form frmTimeKeepingClosed::cmdRecord_Click [8043-8159] [outputTo, runSql]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7416-7497] [runSql]
- form frmTKClose::txtTKButton_Click [2379-2708] [runSql]
- form Time Keeping::cmdRecordShortTK_Click [4237-4354] [outputTo, runSql]
- form Time Keeping::cmdRecordTKStatement_Click [4516-4631] [outputTo, runSql]
- form Time Keeping::cmdCreateAR_Click [4632-4667] [runSql]
- form Time Keeping::cmdInsertTime_Click [4687-4710] [runSql]
- form Time Keeping::cmdAddNew_Click [4711-4732]
- form Time Keeping::addNewTK [4733-4788] [runSql]
- form Time Keeping::cmdCompCurrent_Click [4846-4899]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8029-8053]
- report rpt_Compr_InvoiceADVCur::Report_Open [1528-1564]
- report Invoice2::Charge19_Click [2044-2076] [runSql]
- report Rpt_MergeInvTK::Charge19_Click [1869-1901] [runSql]
- module modGaz::fncRunningDebit [2-17]
- module modGaz::fncRunningCredit [18-33]
- module modGaz::fnc_TEST_get_remaining_AdvancedChargesBalance [399-405]
