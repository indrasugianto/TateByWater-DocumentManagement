# Report Lineage: rptInvoiceComprehensiveAR2

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, tblCase.Retainer, tblCase.CaseOpenDate, [Matter and AR].OrderNr, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Bill_ID FROM (tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID) INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID WHERE ((([Matter and AR].Date2)<=[Bill Closed Date]) AND (([Matter and AR].Charge)<>0)) ORDER BY [Matter and AR].Date2;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Matter and AR, TB Time Keeping

## Related VBA Procedures
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frm_advanced_payments::FilterMe [6918-6944]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389]
- form frmMatter::Date2_AfterUpdate [1512-1562]
- form frmMatter::reorderByDateMatter [1571-1592]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3104-3204]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2573-2682]
- form frmTakeOffSubForm3::CaseNum_Click [2483-2605]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7863-7979]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8000-8042]
- form frmTimeKeepingClosed::cmdRecord_Click [8043-8159]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7416-7497]
- form frmTKClose::txtTKButton_Click [2379-2708]
- form Time Keeping::cmdRecordShortTK_Click [4237-4354]
- form Time Keeping::cmdRecordTKStatement_Click [4516-4631]
- form Time Keeping::cmdCreateAR_Click [4632-4667]
- form Time Keeping::cmdInsertTime_Click [4687-4710]
- form Time Keeping::cmdAddNew_Click [4711-4732]
- form Time Keeping::addNewTK [4733-4788]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8029-8053]
- report Invoice2::Charge19_Click [2044-2076]
- report Rpt_MergeInvTK::Charge19_Click [1869-1901]
