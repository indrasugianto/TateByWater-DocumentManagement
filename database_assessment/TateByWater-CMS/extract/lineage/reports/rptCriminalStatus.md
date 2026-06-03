# Report Lineage: rptCriminalStatus

## Trigger Paths
- frmClientLedger -> CrStatusRep -> onClick -> CrStatusRep_Click (high confidence)
- frmCrimStatusReport -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)
- frmCrimStatusReport -> cmdAllAttyReport -> onClick -> cmdAllAttyReport_Click (high confidence)

## Data Lineage
- RecordSource: `qryCrimStatus`
- RecordSourceType: `saved-query`
- Involved Queries: qryCrimStatus
- Terminal Tables: tblCase, tblDropD

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689] [createObject, runSql]
- form frmClientLedger::CrStatusRep_Click [17058-17077]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmCrimStatusReport::cmdAllAttyReport_Click [673-678]
- form frmCrimStatusReport::cmdAttyReport_Click [679-796]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
