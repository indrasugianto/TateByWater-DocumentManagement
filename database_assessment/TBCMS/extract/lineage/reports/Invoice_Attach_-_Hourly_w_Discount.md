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
- form frmClientLedger::cmdClientReviewEmail_Click [16616-16646] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16647-16677] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17547-17684] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17729-17739]
- form frmTimeKeepingClosed::cmdPrintInvoice_Click [7788-7811]
- form frmTimeKeepingClosed::cmdPreview2_Click [7812-7832]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7863-7979] [outputTo, runSql]
- form frmTimeKeepingClosed::cmdRecord_Click [8043-8159] [outputTo, runSql]
- form frmTimeKeepingOpen::cmdPreview_Click [7350-7371]
- form frmTimeKeepingOpen::cmdPrintInvoice_Click [7372-7395]
- form zfrmSelectCaseNum_Discount::CmdOpenInvoiceAttachReport_Click [400-419]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form Time Keeping::cmdRecordShortTK_Click [4237-4354] [outputTo, runSql]
- form Time Keeping::cmdPrintInvoice_Click [4464-4497]
- form Time Keeping::cmdRecordTKStatement_Click [4516-4631] [outputTo, runSql]
- form Time Keeping::cmdInsertTime_Click [4687-4710] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module DocumentManagement::MoveDocumentByCaseStatus [1373-1585] [createObject, fileSystem]
- module DocumentManagement::GetIntakeDocumentFileName [1682-1758]
- module DocumentManagement::Phase5_E2E_HappyPathTest [1759-2092] [createObject, fileSystem]
- module Module1::GetRetainer [20-23]
