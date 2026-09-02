# Report Lineage: rptBillingTotals

## Trigger Paths
- frm_Billing_Tracker2 -> cmdAttyReport -> onClick -> cmdAttyReport_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT vwBillingTracker2.Tatty, Sum(vwBillingTracker2.Time_) AS SumOfTime_, Sum(vwBillingTracker2.Billed) AS SumOfBilled, [forms]![frm_Billing_Tracker2]![txtFrom] AS StartDate, [forms]![frm_Billing_Tracker2]![txtTo] AS EndDate FROM vwBillingTracker2 WHERE (((vwBillingTracker2.Tdate) Between [forms]![frm_Billing_Tracker2]![txtFrom] And [forms]![frm_Billing_Tracker2]![txtTo])) GROUP BY vwBillingTracker2.Tatty ORDER BY vwBillingTracker2.Tatty DESC;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: vwBillingTracker2

## Related VBA Procedures
- form frm_Billing_Tracker2::cmdAttyReport_Click [6959-6966]
