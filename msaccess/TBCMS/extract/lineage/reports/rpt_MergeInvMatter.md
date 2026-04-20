# Report Lineage: rpt_MergeInvMatter

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, [Matter and AR].Payment, Nz([Charge],0)-Nz([payment],0) AS Balance, fncRunningDebit([tblcase].[CaseID],[Date2],[MatterID]) AS RunningDebit, fncRunningCredit([tblcase].[CaseID],[Date2],[MatterID]) AS RunningCredit, [RunningDebit]-[RunningCredit] AS RunningBalance, tblCase.Retainer, [RunningBalance]+[retainer] AS RetBal FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID ORDER BY [Matter and AR].MatterID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Matter and AR, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form frm_advanced_payments::FilterMe [7008-7035]
- form frmTakeOffReconciliation::txtTKButton_Click [3724-3781]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3192-3292]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmMatter::Date2_AfterUpdate [1513-1563]
- form frmMatter::reorderByDateMatter [1572-1593]
- form Time Keeping::cmdCreateAR_Click [4573-4608]
- form frmTKClose::txtTKButton_Click [2512-2841]
- form frmTakeOffSubForm3::CaseNum_Click [2558-2680]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2653-2762]
- report Invoice2::Charge19_Click [2001-2033]
- report Rpt_MergeInvTK::Charge19_Click [1949-1981]
