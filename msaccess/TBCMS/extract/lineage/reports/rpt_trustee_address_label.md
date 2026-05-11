# Report Lineage: rpt_trustee_address_label

## Trigger Paths
- frmBankruptcy -> cmdPrintTrusteeLabel -> onClick -> cmdPrintTrusteeLabel_Click (high confidence)

## Data Lineage
- RecordSource: `Bankruptcy`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: Bankruptcy

## Related VBA Procedures
- form frmBankruptcy::cmdPrintTrusteeLabel_Click [2447-2462]
- form frmCalls::cmdBankruptcySend_Click [7847-7871]
