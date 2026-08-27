# Report Lineage: rptPersInjuryStatus

## Trigger Paths
- frmPersInjuryStatusReport -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)
- frmPersInjuryStatusReport -> cmdAllAttyReport -> onClick -> cmdAllAttyReport_Click (high confidence)
- frmPersonalInjury -> CrStatusRep -> onClick -> CrStatusRep_Click (high confidence)
- frmOpenReport -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)

## Data Lineage
- RecordSource: `qryPersInjStatus`
- RecordSourceType: `saved-query`
- Involved Queries: qryPersInjStatus
- Terminal Tables: Personal Injury, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmPersInjuryStatusReport::cmdAllAttyReport_Click [575-580]
- form frmPersInjuryStatusReport::cmdAttyReport_Click [581-586]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmPersonalInjury::CrStatusRep_Click [3991-4010]
- form frmOpenReport::cmdAllAttyReport_Click [625-630]
- form frmOpenReport::cmdAttyReport_Click [631-636]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
