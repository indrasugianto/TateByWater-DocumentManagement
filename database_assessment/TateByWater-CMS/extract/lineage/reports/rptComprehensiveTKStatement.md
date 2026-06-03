# Report Lineage: rptComprehensiveTKStatement

## Trigger Paths
- Time Keeping -> cmdTotalTKs -> onClick -> cmdTotalTKs_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceAttachComp`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceAttachComp
- Terminal Tables: tblCase, tblTimeTableDetail

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form Time Keeping::cmdTotalTKs_Click [4384-4398]
- form Time Keeping::cmdInsertTime_Click [4687-4710] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
