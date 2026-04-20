# Report Lineage: Client_Trust_Accounts_for_PreTake_Off

## Trigger Paths
- frmTakeOffReconciliation -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)

## Data Lineage
- RecordSource: `qryTakeOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryTakeOff, qryTakeOff_A
- Terminal Tables: vwTakeOff_A

## Related VBA Procedures
- form frmTakeOffReconciliation::cmdAttyReport_Click [3664-3691]
- form frmTakeOffReconciliation::Form_Load [3692-3715]
- form frmTakeOffReconciliation::cmdInsertData_Click [3863-3942]
- form frmTRUSTENTRIESCHRON::cmdRequery_Click [2643-2652]
- form frmTRUSTENTRIESCHRON::Form_Load [2657-2665]
