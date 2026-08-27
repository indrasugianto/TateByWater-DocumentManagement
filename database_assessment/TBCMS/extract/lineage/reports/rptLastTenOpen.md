# Report Lineage: rptLastTenOpen

## Trigger Paths
- frmCaseListOpen -> cmdLastTenOpen -> onClick -> cmdLastTenOpen_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT qryCaseListOpen.CaseID, qryCaseListOpen.CaseOpenDate, qryCaseListOpen.ClientName, qryCaseListOpen.Orig_Atty, qryCaseListOpen.Matter_type, qryCaseListOpen.FileNumber, qryCaseListOpen.Retainer, qryCaseListOpen.Number_, qryCaseListOpen.yr, qryCaseListOpen.Case_Letter, tblCase.Referral, qryCaseListOpen.CodeVal FROM qryCaseListOpen INNER JOIN tblCase ON qryCaseListOpen.CaseID = tblCase.CaseID WHERE (((qryCaseListOpen.CaseOpenDate) Between getSTDT() And getENDT()));`
- RecordSourceType: `inline-sql`
- Involved Queries: qryCaseListOpen
- Terminal Tables: tblCase, vwCaseListOpen

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmCaseListOpen::cmdLastTenOpen_Click [1549-1563]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
