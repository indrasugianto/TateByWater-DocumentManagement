# Report Lineage: rpt_Compr_InvoiceADVCur

## Trigger Paths
- frmTimeKeepingClosed -> cmdCompCurr -> onClick -> cmdCompCurr_Click (high confidence)
- Time Keeping -> cmdCompCurrent -> onClick -> cmdCompCurrent_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: TB Time Keeping, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmTakeOffReconciliation::txtTKButton_Click [3332-3389]
- form frmTimeKeepingClosed::cmdCompCurr_Click [7680-7728]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7863-7979]
- form frmTimeKeepingClosed::cmdAddNewTK_Click [8000-8042]
- form frmTimeKeepingClosed::cmdRecord_Click [8043-8159]
- form frmTimeKeepingOpen::cmdAddNewTK_Click [7416-7497]
- form frmTKClose::txtTKButton_Click [2379-2708]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form Time Keeping::cmdRecordShortTK_Click [4237-4354]
- form Time Keeping::cmdRecordTKStatement_Click [4516-4631]
- form Time Keeping::cmdCreateAR_Click [4632-4667]
- form Time Keeping::cmdInsertTime_Click [4687-4710]
- form Time Keeping::cmdAddNew_Click [4711-4732]
- form Time Keeping::addNewTK [4733-4788]
- form Time Keeping::cmdCompCurrent_Click [4846-4899]
- form zClient Ledger OLD::cmdCreateHourlyBill_Click [8029-8053]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
