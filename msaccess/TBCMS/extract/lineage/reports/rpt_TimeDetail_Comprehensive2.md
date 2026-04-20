# Report Lineage: rpt_TimeDetail_Comprehensive2

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `qryInvoiceComprehensiveTimeDetail2`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceComprehensiveTimeDetail2, qryInvoiceAttachRPT
- Terminal Tables: tblCase, tblTimeTableDetail

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form Time Keeping::cmdInsertTime_Click [4628-4651]
