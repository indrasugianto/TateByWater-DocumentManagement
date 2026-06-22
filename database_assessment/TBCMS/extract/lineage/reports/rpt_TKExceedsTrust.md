# Report Lineage: rpt_TKExceedsTrust

## Trigger Paths
- frmTimeKeepingClosed -> cmdPreview -> onClick -> cmdPreview_Click (high confidence)
- Time Keeping -> cmdPrevStatement -> onClick -> cmdPrevStatement_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceAttachRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceAttachRPT1, qryInvoiceAttachRPT
- Terminal Tables: TblCase, tblCase, tblTimeTableDetail

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmTimeKeepingClosed::cmdPreview_Click [7762-7787]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form Time Keeping::cmdPrevStatement_Click [4204-4233]
- form Time Keeping::cmdInsertTime_Click [4687-4710]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
