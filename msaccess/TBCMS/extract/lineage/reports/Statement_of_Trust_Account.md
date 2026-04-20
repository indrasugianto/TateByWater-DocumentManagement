# Report Lineage: Statement of Trust Account

## Trigger Paths
- frm_trust_summary -> cmdPrintStmtTrust -> onClick -> cmdPrintStmtTrust_Click (high confidence)
- frm_trust_summary -> cmdPreviewTrustStmt -> onClick -> cmdPreviewTrustStmt_Click (high confidence)
- frmClientLedger -> cmdAddNew -> onClick -> cmdAddNew_Click (high confidence)
- frmClientLedger -> cmdStatementTrustAcct -> onClick -> cmdStatementTrustAcct_Click (high confidence)
- zClient Ledger OLD -> cmdStatemenTrustAccount -> onClick -> cmdStatemenTrustAccount_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)

## Data Lineage
- RecordSource: `qryStmtTrustRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryStmtTrustRPT1
- Terminal Tables: vwStmtTrustRPT1

## Related VBA Procedures
- form frm_trust_summary::cmdPreviewTrustStmt_Click [6942-6968]
- form frm_trust_summary::cmdPrintStmtTrust_Click [6969-6984]
- form frmClientLedger::cmdStatementTrustAcct_Click [16820-16832]
- form frmClientLedger::cmdAddNew_Click [17143-17209]
- form zClient Ledger OLD::cmdStatemenTrustAccount_Click [7905-7917]
- form frmHome::cmdDeleteAll_Click [1920-2057]
- form frmHomeAdmin::cmdDeleteAll_Click [6356-6493]
