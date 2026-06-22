# Report Lineage: rpt_Disposition_Closing

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `qry_Disposition_ClosingSheet`
- RecordSourceType: `saved-query`
- Involved Queries: qry_Disposition_ClosingSheet
- Terminal Tables: Disposition, TblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmDispositions::FilterMe [7360-7390]
- form frmHomeAdmin::cmdDispositions_Click [6225-6228]
- form frmHomeAdmin::Command119_Click [6469-6472]
- form frmHome::cmdDispositions_Click [1800-1803]
- form frmHome::Command119_Click [2024-2027]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
