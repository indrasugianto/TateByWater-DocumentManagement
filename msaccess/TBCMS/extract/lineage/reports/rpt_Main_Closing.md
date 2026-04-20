# Report Lineage: rpt_Main_Closing

## Trigger Paths
- frmClientLedger -> Command353 -> onClick -> Command353_Click (high confidence)
- frmClientLedger -> cmdClosingSheet -> onClick -> cmdClosingSheet_Click (high confidence)
- zClient Ledger OLD -> Command353 -> onClick -> Command353_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmToBeClosed -> Command353 -> onClick -> Command353_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmdClosingSheet_Click [16544-16555]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form frmClientLedger::Command353_Click [17559-17569]
- form zClient Ledger OLD::Command353_Click [7734-7740]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmHome::cmdDeleteAll_Click [1920-2057]
- form frmToBeClosed::Command353_Click [6856-6863]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmHomeAdmin::cmdDeleteAll_Click [6356-6493]
