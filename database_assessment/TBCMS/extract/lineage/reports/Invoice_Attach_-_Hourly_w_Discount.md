# Report Lineage: Invoice Attach - Hourly w Discount

## Trigger Paths
- frmTimeKeepingOpen -> cmdPrintInvoice -> onClick -> cmdPrintInvoice_Click (high confidence)
- frmTimeKeepingOpen -> cmdPreview -> onClick -> cmdPreview_Click (high confidence)
- zfrmSelectCaseNum_Discount -> CmdOpenInvoiceAttachReport -> onClick -> CmdOpenInvoiceAttachReport_Click (medium confidence)
- frmTimeKeepingClosed -> cmdRecord -> onClick -> cmdRecord_Click (high confidence)
- frmTimeKeepingClosed -> cmdPreview2 -> onClick -> cmdPreview2_Click (high confidence)
- frmTimeKeepingClosed -> cmdRecordShort -> onClick -> cmdRecordShort_Click (high confidence)
- Time Keeping -> cmdPrintInvoice -> onClick -> cmdPrintInvoice_Click (high confidence)
- Time Keeping -> cmdRecordTKStatement -> onClick -> cmdRecordTKStatement_Click (high confidence)
- Time Keeping -> cmdRecordShortTK -> onClick -> cmdRecordShortTK_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceAttachRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceAttachRPT1, qryInvoiceAttachRPT
- Terminal Tables: TblCase, tblCase, tblTimeTableDetail

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form frmTimeKeepingOpen::cmdPreview_Click [7386-7407]
- form frmTimeKeepingOpen::cmdPrintInvoice_Click [7408-7431]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form zfrmSelectCaseNum_Discount::CmdOpenInvoiceAttachReport_Click [391-410]
- form frmTimeKeepingClosed::cmdPrintInvoice_Click [7850-7873]
- form frmTimeKeepingClosed::cmdPreview2_Click [7874-7894]
- form frmTimeKeepingClosed::cmdRecordShort_Click [7925-8041] [outputTo, runSql]
- form frmTimeKeepingClosed::cmdRecord_Click [8105-8221] [outputTo, runSql]
- form Time Keeping::cmdRecordShortTK_Click [4238-4355] [outputTo, runSql]
- form Time Keeping::cmdPrintInvoice_Click [4465-4498]
- form Time Keeping::cmdRecordTKStatement_Click [4517-4632] [outputTo, runSql]
- form Time Keeping::cmdInsertTime_Click [4688-4711] [runSql]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
