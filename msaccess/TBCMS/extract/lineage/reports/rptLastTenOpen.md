# Report Lineage: rptLastTenOpen

## Trigger Paths
- frmCaseListOpen -> cmdLastTenOpen -> onClick -> cmdLastTenOpen_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT qryCaseListOpen.CaseID, qryCaseListOpen.CaseOpenDate, qryCaseListOpen.ClientName, qryCaseListOpen.Orig_Atty, qryCaseListOpen.Matter_type, qryCaseListOpen.FileNumber, qryCaseListOpen.Retainer, qryCaseListOpen.Number_, qryCaseListOpen.yr, qryCaseListOpen.Case_Letter, tblCase.Referral, qryCaseListOpen.CodeVal FROM qryCaseListOpen INNER JOIN tblCase ON qryCaseListOpen.CaseID = tblCase.CaseID WHERE (((qryCaseListOpen.CaseOpenDate) Between getSTDT() And getENDT()));`
- RecordSourceType: `inline-sql`
- Involved Queries: qryCaseListOpen
- Terminal Tables: tblCase, vwCaseListOpen

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmCaseListOpen::cmdLastTenOpen_Click [1549-1563]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
