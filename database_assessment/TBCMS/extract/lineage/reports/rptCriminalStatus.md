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
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::CrStatusRep_Click [18862-18881]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmCrimStatusReport::cmdAllAttyReport_Click [674-679]
- form frmCrimStatusReport::cmdAttyReport_Click [680-797]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
