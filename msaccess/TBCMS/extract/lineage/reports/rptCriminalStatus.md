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
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::CrStatusRep_Click [17058-17077]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmCrimStatusReport::cmdAllAttyReport_Click [673-678]
- form frmCrimStatusReport::cmdAttyReport_Click [679-796]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
