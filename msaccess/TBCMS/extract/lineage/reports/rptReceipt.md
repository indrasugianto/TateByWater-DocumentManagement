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
- form frmMatter::cmdReceipt_Click [1431-1453]
- form frmMatter::Date2_AfterUpdate [1513-1563]
- form frmMatter::reorderByDateMatter [1572-1593]
- form Time Keeping::cmdCreateAR_Click [4573-4608]
- form frmTKClose::txtTKButton_Click [2512-2841]
- form frmReceipt::cmdGenerateReceipt_Click [1219-1246]
- form frmReceipt::Command26_Click [1247-1254]
- form frmTakeOffSubForm3::CaseNum_Click [2558-2680]
- form frmTakeOffSubForm_OLD::cmdInsertIntoTA_Click [2653-2762]
- report Invoice2::Charge19_Click [2001-2033]
- report Rpt_MergeInvTK::Charge19_Click [1949-1981]
