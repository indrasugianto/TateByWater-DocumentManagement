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
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmAttyNotes::cmdPrintNotes_Click [612-633]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmUpcoming Hearings::cmdPrintNotes_Click [7565-7584]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
