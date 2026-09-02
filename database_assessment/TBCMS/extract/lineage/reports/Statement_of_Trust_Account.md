# Report Lineage: Statement of Trust Account

## Trigger Paths
- frmClientLedger -> cmdAddNew -> onClick -> cmdAddNew_Click (high confidence)
- frmClientLedger -> cmdStatementTrustAcct -> onClick -> cmdStatementTrustAcct_Click (high confidence)
- frm_trust_summary -> cmdPrintStmtTrust -> onClick -> cmdPrintStmtTrust_Click (high confidence)
- frm_trust_summary -> cmdPreviewTrustStmt -> onClick -> cmdPreviewTrustStmt_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- zClient Ledger OLD -> cmdStatemenTrustAccount -> onClick -> cmdStatemenTrustAccount_Click (high confidence)

## Data Lineage
- RecordSource: `qryStmtTrustRPT1`
- RecordSourceType: `saved-query`
- Involved Queries: qryStmtTrustRPT1
- Terminal Tables: vwStmtTrustRPT1

## Related VBA Procedures
- form frmClientLedger::cmdStatementTrustAcct_Click [16989-17001]
- form frmClientLedger::cmdAddNew_Click [17332-17398]
- form frm_trust_summary::cmdPreviewTrustStmt_Click [6922-6948]
- form frm_trust_summary::cmdPrintStmtTrust_Click [6949-6964]
- form frmHomeAdmin::cmdDeleteAll_Click [6313-6450] [runSql]
- form frmHome::cmdDeleteAll_Click [1832-1969] [runSql]
- form zClient Ledger OLD::cmdStatemenTrustAccount_Click [7926-7938]
