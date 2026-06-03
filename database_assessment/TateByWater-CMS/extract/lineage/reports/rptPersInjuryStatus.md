# Report Lineage: rptPersInjuryStatus

## Trigger Paths
- frmPersInjuryStatusReport -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)
- frmPersInjuryStatusReport -> cmdAllAttyReport -> onClick -> cmdAllAttyReport_Click (high confidence)
- frmOpenReport -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)
- frmPersonalInjury -> CrStatusRep -> onClick -> CrStatusRep_Click (high confidence)

## Data Lineage
- RecordSource: `qryPersInjStatus`
- RecordSourceType: `saved-query`
- Involved Queries: qryPersInjStatus
- Terminal Tables: Personal Injury, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmPersInjuryStatusReport::cmdAllAttyReport_Click [583-588]
- form frmPersInjuryStatusReport::cmdAttyReport_Click [589-594]
- form frmOpenReport::cmdAllAttyReport_Click [625-630]
- form frmOpenReport::cmdAttyReport_Click [631-636]
- form frmPersonalInjury::CrStatusRep_Click [3818-3837]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
