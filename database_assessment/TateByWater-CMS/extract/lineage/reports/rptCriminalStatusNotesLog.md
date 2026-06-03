# Report Lineage: rptCriminalStatusNotesLog

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `SELECT tblNotes.CaseID, tblNotes.NoteDate, tblNotes.NotePerson, tblNotes.NoteDescription, tblNotes.NoteTime, tblNotes.CaseID, tblNotes.IDNotes FROM tblNotes;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: tblNotes

## Related VBA Procedures
- No related VBA procedure could be inferred.
