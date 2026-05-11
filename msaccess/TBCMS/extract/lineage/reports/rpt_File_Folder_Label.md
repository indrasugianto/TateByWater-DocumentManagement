# Report Lineage: rpt_File_Folder_Label

## Trigger Paths
- frmClientLedger -> cmdPrintFileLabel -> onClick -> cmdPrintFileLabel_Click (high confidence)

## Data Lineage
- RecordSource: `qryFileFolderLabel`
- RecordSourceType: `saved-query`
- Involved Queries: qryFileFolderLabel
- Terminal Tables: tblCase, tblHearingDate

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmdPrintFileLabel_Click [16880-16895]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
