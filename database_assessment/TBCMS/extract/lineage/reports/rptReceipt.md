# Report Lineage: rptReceipt

## Trigger Paths
- frmMatter -> cmdReceipt -> onClick -> cmdReceipt_Click (high confidence)
- frmReceipt -> cmdGenerateReceipt -> onClick -> cmdGenerateReceipt_Click (medium confidence)

## Data Lineage
- RecordSource: `qryReceipt`
- RecordSourceType: `saved-query`
- Involved Queries: qryReceipt
- Terminal Tables: Matter and AR, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form frm_advanced_payments::FilterMe [6919-6945]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmTakeOffSubForm::cmdInsertIntoTA_Click [3103-3203] [runSql]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmMatter::cmdReceipt_Click [1431-1453]
- form frmMatter::Date2_AfterUpdate [1513-1563] [runSql]
- form frmMatter::reorderByDateMatter [1572-1593] [runSql]
- form Time Keeping::cmdCreateAR_Click [4633-4668] [runSql]
- form frmTKClose::txtTKButton_Click [2378-2707] [runSql]
- form frmReceipt::cmdGenerateReceipt_Click [1219-1246]
- form frmReceipt::Command26_Click [1247-1254]
- form frmTakeOffSubForm3::CaseNum_Click [2558-2680] [runSql]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2653-2762] [runSql]
- report Invoice2::Charge19_Click [2068-2100] [runSql]
- report Rpt_MergeInvTK::Charge19_Click [1949-1981] [runSql]
- module modGaz::fncRunningDebit [2-17]
- module modGaz::fncRunningCredit [18-33]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fnc_TEST_get_remaining_AdvancedChargesBalance [399-405]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
