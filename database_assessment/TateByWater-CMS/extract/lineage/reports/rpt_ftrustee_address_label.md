# Report Lineage: rpt_ftrustee_address_label

## Trigger Paths
- frmBankruptcy -> cmdPrintForeLabel -> onClick -> cmdPrintForeLabel_Click (high confidence)

## Data Lineage
- RecordSource: `Bankruptcy`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: Bankruptcy

## Related VBA Procedures
- form frmBankruptcy::cmdPrintForeLabel_Click [2431-2446]
- form frmCalls::cmdBankruptcySend_Click [7847-7871] [createObject]
