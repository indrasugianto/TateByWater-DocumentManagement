# Report Lineage: rpt_Open_Cases

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `qryTakeOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryTakeOff, qryTakeOff_A
- Terminal Tables: vwTakeOff_A

## Related VBA Procedures
- form frmTakeOffReconciliation::Form_Load [3300-3323]
- form frmTakeOffReconciliation::cmdInsertData_Click [3471-3550] [runSql]
- form frmTRUSTENTRIESCHRON::cmdRequery_Click [2480-2489]
- form frmTRUSTENTRIESCHRON::Form_Load [2494-2502]
