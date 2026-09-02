# Report Lineage: Client_Trust_Accounts_for_PreTake_Off

## Trigger Paths
- frmTakeOffReconciliation -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)

## Data Lineage
- RecordSource: `qryTakeOff`
- RecordSourceType: `saved-query`
- Involved Queries: qryTakeOff, qryTakeOff_A
- Terminal Tables: vwTakeOff_A

## Related VBA Procedures
- form frmTakeOffReconciliation::cmdAttyReport_Click [3265-3317] [runSql]
- form frmTakeOffReconciliation::Form_Load [3318-3347]
- form frmTakeOffReconciliation::cmdInsertData_Click [3495-3574] [runSql]
- form frmTRUSTENTRIESCHRON::FilterClear [2448-2483]
- form frmTRUSTENTRIESCHRON::cmdRequery_Click [2484-2499]
- form frmTRUSTENTRIESCHRON::cmdWFReconcile_AfterUpdate [2500-2514]
- form frmTRUSTENTRIESCHRON::Form_Load [2515-2529]
