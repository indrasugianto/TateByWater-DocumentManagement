# Report Lineage: rpt_Main_Closing

## Trigger Paths
- frmClientLedger -> Command353 -> onClick -> Command353_Click (high confidence)
- frmClientLedger -> cmdClosingSheet -> onClick -> cmdClosingSheet_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmToBeClosed -> Command353 -> onClick -> Command353_Click (high confidence)
- zClient Ledger OLD -> Command353 -> onClick -> Command353_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmdClosingSheet_Click [16725-16736]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmClientLedger::Command353_Click [17760-17770]
- form frmHomeAdmin::cmdDeleteAll_Click [6313-6450]
- form frmHome::cmdDeleteAll_Click [1868-2005]
- form frmToBeClosed::Command353_Click [6821-6828]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::Command353_Click [7755-7761]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
