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
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmPersInjuryStatusReport::cmdAllAttyReport_Click [575-580]
- form frmPersInjuryStatusReport::cmdAttyReport_Click [581-586]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmPersonalInjury::CrStatusRep_Click [3991-4010]
- form frmOpenReport::cmdAllAttyReport_Click [625-630]
- form frmOpenReport::cmdAttyReport_Click [631-636]
