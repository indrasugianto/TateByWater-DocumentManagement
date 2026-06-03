# Report Lineage: rptReceipt

## Trigger Paths
- frmReceipt -> cmdGenerateReceipt -> onClick -> cmdGenerateReceipt_Click (medium confidence)
- frmMatter -> cmdReceipt -> onClick -> cmdReceipt_Click (high confidence)

## Data Lineage
- RecordSource: `qryReceipt`
- RecordSourceType: `saved-query`
- Involved Queries: qryReceipt
- Terminal Tables: Matter and AR, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frm_advanced_payments::FilterMe [6918-6944]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389] [runSql]
- form frmReceipt::cmdGenerateReceipt_Click [1205-1232]
- form frmReceipt::Command26_Click [1233-1240]
- form frmMatter::cmdReceipt_Click [1430-1452]
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
- module Module1::GetRetainer [20-23]
