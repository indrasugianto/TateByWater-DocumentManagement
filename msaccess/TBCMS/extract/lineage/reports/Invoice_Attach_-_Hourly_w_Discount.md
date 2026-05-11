# Report Lineage: Invoice Attach - Hourly w Discount

## Trigger Paths
- frmTimeKeepingClosed -> cmdRecord -> onClick -> cmdRecord_Click (high confidence)
- frmTimeKeepingClosed -> cmdPreview2 -> onClick -> cmdPreview2_Click (high confidence)
- frmTimeKeepingClosed -> cmdRecordShort -> onClick -> cmdRecordShort_Click (high confidence)
- frmTimeKeepingOpen -> cmdPrintInvoice -> onClick -> cmdPrintInvoice_Click (high confidence)
- frmTimeKeepingOpen -> cmdPreview -> onClick -> cmdPreview_Click (high confidence)
- zfrmSelectCaseNum_Discount -> CmdOpenInvoiceAttachReport -> onClick -> CmdOpenInvoiceAttachReport_Click (medium confidence)
- Time Keeping -> cmdPrintInvoice -> onClick -> cmdPrintInvoice_Click (high confidence)
- Time Keeping -> cmdRecordTKStatement -> onClick -> cmdRecordTKStatement_Click (high confidence)
- Time Keeping -> cmdRecordShortTK -> onClick -> cmdRecordShortTK_Click (high confidence)

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
- form frmTimeKeepingClosed::cmdPrintInvoice_Click [7788-7811]
- form frmTimeKeepingClosed::cmdPreview2_Click [7812-7832]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7863-7979]
- form frmTimeKeepingClosed::cmdRecord_Click [8043-8159]
- form frmTimeKeepingOpen::cmdPreview_Click [7350-7371]
- form frmTimeKeepingOpen::cmdPrintInvoice_Click [7372-7395]
- form zfrmSelectCaseNum_Discount::CmdOpenInvoiceAttachReport_Click [400-419]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form Time Keeping::cmdRecordShortTK_Click [4237-4354]
- form Time Keeping::cmdPrintInvoice_Click [4464-4497]
- form Time Keeping::cmdRecordTKStatement_Click [4516-4631]
- form Time Keeping::cmdInsertTime_Click [4687-4710]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
