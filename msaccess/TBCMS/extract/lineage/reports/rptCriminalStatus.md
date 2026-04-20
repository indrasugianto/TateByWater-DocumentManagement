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
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::CrStatusRep_Click [16857-16876]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmCrimStatusReport::cmdAllAttyReport_Click [674-679]
- form frmCrimStatusReport::cmdAttyReport_Click [680-797]
