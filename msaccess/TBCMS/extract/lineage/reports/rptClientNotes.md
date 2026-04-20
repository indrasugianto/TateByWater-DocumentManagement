# Report Lineage: rptClientNotes

## Trigger Paths
- frmAttyNotes -> cmdPrintNotes -> onClick -> cmdPrintNotes_Click (high confidence)
- frmUpcoming Hearings -> cmdPrintNotes -> onClick -> cmdPrintNotes_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblNotes.NoteDate, tblNotes.NotePerson, tblNotes.NoteDescription, tblNotes.NoteTime, Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNo, tblNotes.IDNotes FROM tblCase INNER JOIN tblNotes ON tblCase.CaseID = tblNotes.CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: tblCase, tblNotes

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmAttyNotes::cmdPrintNotes_Click [612-633]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmUpcoming Hearings::cmdPrintNotes_Click [7565-7584]
