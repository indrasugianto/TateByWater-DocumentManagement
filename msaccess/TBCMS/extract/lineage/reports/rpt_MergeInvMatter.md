# Report Lineage: rpt_MergeInvMatter

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, [Matter and AR].Payment, Nz([Charge],0)-Nz([payment],0) AS Balance, fncRunningDebit([tblcase].[CaseID],[Date2],[MatterID]) AS RunningDebit, fncRunningCredit([tblcase].[CaseID],[Date2],[MatterID]) AS RunningCredit, [RunningDebit]-[RunningCredit] AS RunningBalance, tblCase.Retainer, [RunningBalance]+[retainer] AS RetBal FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID ORDER BY [Matter and AR].MatterID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Matter and AR, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frm_advanced_payments::FilterMe [6918-6944]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389]
- form frmMatter::Date2_AfterUpdate [1512-1562]
- form frmMatter::reorderByDateMatter [1571-1592]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3104-3204]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2573-2682]
- form frmTakeOffSubForm3::CaseNum_Click [2483-2605]
- form frmTKClose::txtTKButton_Click [2379-2708]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form Time Keeping::cmdCreateAR_Click [4632-4667]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- report Invoice2::Charge19_Click [2044-2076]
- report Rpt_MergeInvTK::Charge19_Click [1869-1901]
