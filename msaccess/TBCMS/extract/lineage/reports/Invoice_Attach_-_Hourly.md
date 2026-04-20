# Report Lineage: Invoice Attach - Hourly

## Trigger Paths
- frmTimeKeepingOpen -> cmdPrintInvoice -> onClick -> cmdPrintInvoice_Click (high confidence)
- frmTimeKeepingOpen -> cmdPreview -> onClick -> cmdPreview_Click (high confidence)
- zfrmSelectCaseNum -> CmdOpenInvoiceAttachReport -> onClick -> CmdOpenInvoiceAttachReport_Click (medium confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- zfrmSelectCaseNum_Discount -> CmdOpenInvoiceAttachReport -> onClick -> CmdOpenInvoiceAttachReport_Click (medium confidence)
- frmTimeKeepingClosed -> cmdRecord -> onClick -> cmdRecord_Click (high confidence)
- frmTimeKeepingClosed -> cmdPreview2 -> onClick -> cmdPreview2_Click (high confidence)
- frmTimeKeepingClosed -> cmdRecordShort -> onClick -> cmdRecordShort_Click (high confidence)
- Time Keeping -> cmdPrintInvoice -> onClick -> cmdPrintInvoice_Click (high confidence)
- Time Keeping -> cmdPreview -> onClick -> cmdPreview_Click (high confidence)
- Time Keeping -> cmdRecordTKStatement -> onClick -> cmdRecordTKStatement_Click (high confidence)
- Time Keeping -> cmdRecordShortTK -> onClick -> cmdRecordShortTK_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceAttachRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceAttachRPT1, qryInvoiceAttachRPT
- Terminal Tables: TblCase, tblCase, tblTimeTableDetail

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form frmTimeKeepingOpen::cmdPreview_Click [7386-7407]
- form frmTimeKeepingOpen::cmdPrintInvoice_Click [7408-7431]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form zfrmSelectCaseNum::CmdOpenInvoiceAttachReport_Click [385-404]
- form frmHome::cmdDeleteAll_Click [1920-2057]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form zfrmSelectCaseNum_Discount::CmdOpenInvoiceAttachReport_Click [391-410]
- form frmTimeKeepingClosed::cmdPrintInvoice_Click [7850-7873]
- form frmTimeKeepingClosed::cmdPreview2_Click [7874-7894]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7925-8041]
- form frmTimeKeepingClosed::cmdRecord_Click [8105-8221]
- form Time Keeping::cmdRecordShortTK_Click [4199-4316]
- form Time Keeping::cmdPrintInvoice_Click [4405-4438]
- form Time Keeping::cmdRecordTKStatement_Click [4457-4572]
- form Time Keeping::cmdPreview_Click [4609-4627]
- form Time Keeping::cmdInsertTime_Click [4628-4651]
- form frmHomeAdmin::cmdDeleteAll_Click [6356-6493]
