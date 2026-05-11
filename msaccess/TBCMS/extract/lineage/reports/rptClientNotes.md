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
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmAttyNotes::cmdPrintNotes_Click [616-637]
- form frmUpcoming Hearings::cmdPrintNotes_Click [7504-7523]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
