# Report Lineage: rpt_Matter_Closing

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `SELECT vw_rpt_Matter_Closing.CaseID, vw_rpt_Matter_Closing.MatterID, vw_rpt_Matter_Closing.Date2, vw_rpt_Matter_Closing.Pay_Outlay, vw_rpt_Matter_Closing.Charge, vw_rpt_Matter_Closing.Payment, vw_rpt_Matter_Closing.Balance, vw_rpt_Matter_Closing.RunningDebit, vw_rpt_Matter_Closing.RunningCredit, vw_rpt_Matter_Closing.RunningBalance, vw_rpt_Matter_Closing.Retainer, vw_rpt_Matter_Closing.RetBal, vw_rpt_Matter_Closing.OrderNr FROM vw_rpt_Matter_Closing ORDER BY vw_rpt_Matter_Closing.MatterID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: vw_rpt_Matter_Closing

## Related VBA Procedures
- No related VBA procedure could be inferred.
