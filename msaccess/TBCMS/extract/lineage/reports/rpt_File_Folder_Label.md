# Report Lineage: rpt_File_Folder_Label

## Trigger Paths
- frmClientLedger -> cmdPrintFileLabel -> onClick -> cmdPrintFileLabel_Click (high confidence)

## Data Lineage
- RecordSource: `qryFileFolderLabel`
- RecordSourceType: `saved-query`
- Involved Queries: qryFileFolderLabel
- Terminal Tables: tblCase, tblHearingDate

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmdPrintFileLabel_Click [16699-16714]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
