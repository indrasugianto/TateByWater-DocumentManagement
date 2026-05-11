# Report Lineage: Client_Trust_Accounts_for_PreTake_Off

## Trigger Paths
- frmTakeOffReconciliation -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)

## Data Lineage
- RecordSource: `qryTakeOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryTakeOff, qryTakeOff_A
- Terminal Tables: vwTakeOff_A

## Related VBA Procedures
- form frmTakeOffReconciliation::cmdAttyReport_Click [3272-3299]
- form frmTakeOffReconciliation::Form_Load [3300-3323]
- form frmTakeOffReconciliation::cmdInsertData_Click [3471-3550]
- form frmTRUSTENTRIESCHRON::cmdRequery_Click [2480-2489]
- form frmTRUSTENTRIESCHRON::Form_Load [2494-2502]
