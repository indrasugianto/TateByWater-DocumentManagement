# Report Lineage: rptInvoiceComprehensiveAR

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, [Matter and AR].Payment, tblCase.Retainer, tblCase.CaseOpenDate FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID ORDER BY [Matter and AR].Date2;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Matter and AR, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16616-16646] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16647-16677] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17547-17684] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17729-17739]
- form frm_advanced_payments::FilterMe [6918-6944]
- form frmTakeOffReconciliation::txtTKButton_Click [3356-3413] [runSql]
- form frmMatter::Date2_AfterUpdate [1512-1562] [runSql]
- form frmMatter::reorderByDateMatter [1571-1592] [runSql]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3104-3204] [runSql]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2573-2682] [runSql]
- form frmTakeOffSubForm3::CaseNum_Click [2483-2605] [runSql]
- form frmTKClose::txtTKButton_Click [2379-2708] [runSql]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form Time Keeping::cmdCreateAR_Click [4632-4667] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- report Invoice2::Charge19_Click [2044-2076] [runSql]
- report Rpt_MergeInvTK::Charge19_Click [1869-1901] [runSql]
- module modGaz::fncRunningDebit [2-17]
- module modGaz::fncRunningCredit [18-33]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fnc_TEST_get_remaining_AdvancedChargesBalance [399-405]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module DocumentManagement::MoveDocumentByCaseStatus [1373-1585] [createObject, fileSystem]
- module DocumentManagement::GetIntakeDocumentFileName [1682-1758]
- module DocumentManagement::Phase5_E2E_HappyPathTest [1759-2092] [createObject, fileSystem]
- module Module1::GetRetainer [20-23]
