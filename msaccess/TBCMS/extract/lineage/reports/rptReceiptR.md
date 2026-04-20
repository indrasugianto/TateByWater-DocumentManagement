# Report Lineage: rptReceiptR

## Trigger Paths
- frmReceipt -> cmdGenerateReceipt -> onClick -> cmdGenerateReceipt_Click (high confidence)

## Data Lineage
- RecordSource: `tblReceipts`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblReceipts

## Related VBA Procedures
- form frmReceipt::cmdGenerateReceipt_Click [1219-1246]
- form frmReceipt::Command26_Click [1247-1254]
