# Report Lineage: rpt_OpenCases

## Trigger Paths
- frmCaseListOpen -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)

## Data Lineage
- RecordSource: `qryCaseListOpen`
- RecordSourceType: `saved-query`
- Involved Queries: qryCaseListOpen
- Terminal Tables: vwCaseListOpen

## Related VBA Procedures
- form frmCaseListOpen::cmdAttyReport_Click [1540-1548]
