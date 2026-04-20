# Report Lineage: rpt_Open_Cases

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `qryTakeOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryTakeOff, qryTakeOff_A
- Terminal Tables: vwTakeOff_A

## Related VBA Procedures
- form frmTakeOffReconciliation::Form_Load [3692-3715]
- form frmTakeOffReconciliation::cmdInsertData_Click [3863-3942]
- form frmTRUSTENTRIESCHRON::cmdRequery_Click [2643-2652]
- form frmTRUSTENTRIESCHRON::Form_Load [2657-2665]
