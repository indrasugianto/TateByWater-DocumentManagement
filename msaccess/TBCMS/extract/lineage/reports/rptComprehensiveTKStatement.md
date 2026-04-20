# Report Lineage: rptComprehensiveTKStatement

## Trigger Paths
- Time Keeping -> cmdTotalTKs -> onClick -> cmdTotalTKs_Click (high confidence)

## Data Lineage
- RecordSource: `qryInvoiceAttachComp`
- RecordSourceType: `saved-query`
- Involved Queries: qryInvoiceAttachComp
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
- form Time Keeping::cmdTotalTKs_Click [4325-4339]
- form Time Keeping::cmdInsertTime_Click [4628-4651]
