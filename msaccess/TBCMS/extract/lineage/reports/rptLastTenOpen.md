# Report Lineage: rptLastTenOpen

## Trigger Paths
- frmCaseListOpen -> cmdLastTenOpen -> onClick -> cmdLastTenOpen_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT qryCaseListOpen.CaseID, qryCaseListOpen.CaseOpenDate, qryCaseListOpen.ClientName, qryCaseListOpen.Orig_Atty, qryCaseListOpen.Matter_type, qryCaseListOpen.FileNumber, qryCaseListOpen.Retainer, qryCaseListOpen.Number_, qryCaseListOpen.yr, qryCaseListOpen.Case_Letter, tblCase.Referral, qryCaseListOpen.CodeVal FROM qryCaseListOpen INNER JOIN tblCase ON qryCaseListOpen.CaseID = tblCase.CaseID WHERE (((qryCaseListOpen.CaseOpenDate) Between getSTDT() And getENDT()));`
- RecordSourceType: `inline-sql`
- Involved Queries: qryCaseListOpen
- Terminal Tables: tblCase, vwCaseListOpen

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmCaseListOpen::cmdLastTenOpen_Click [1488-1502]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
