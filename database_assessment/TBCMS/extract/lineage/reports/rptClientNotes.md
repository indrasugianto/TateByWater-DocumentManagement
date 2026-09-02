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
- form frmClientLedger::cmdClientReviewEmail_Click [16616-16646] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16647-16677] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17547-17684] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17729-17739]
- form frmAttyNotes::cmdPrintNotes_Click [616-637]
- form frmUpcoming Hearings::cmdPrintNotes_Click [7504-7523]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module DocumentManagement::MoveDocumentByCaseStatus [1373-1585] [createObject, fileSystem]
- module DocumentManagement::GetIntakeDocumentFileName [1682-1758]
- module DocumentManagement::Phase5_E2E_HappyPathTest [1759-2092] [createObject, fileSystem]
- module Module1::GetRetainer [20-23]
