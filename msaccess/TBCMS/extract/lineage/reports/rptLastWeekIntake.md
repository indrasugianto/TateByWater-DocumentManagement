# Report Lineage: rptLastWeekIntake

## Trigger Paths
- Intakes -> cmdLastWeekIntakes -> onClick -> cmdLastWeekIntakes_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT [TB Intakes].ID, [TB Intakes].[GI Last Name], [TB Intakes].[GI First Name], [TB Intakes].[GI phone], [TB Intakes].[GI Date], [TB Intakes].[GI Practice Area], [TB Intakes].[GI Individual Referrer], [TB Intakes].[GI Comments], [TB Intakes].[GI No Further Action], [TB Intakes].[GI Open], [TB Intakes].[GI Open Date], [TB Intakes].[GI Referral], [TB Intakes].ReasonDintHire, [TB Intakes].FollowUpDate, [TB Intakes].Attorny, [TB Intakes].QuotedFee FROM [TB Intakes] WHERE ((([TB Intakes].[GI Date]) Between getSTDT() And getENDT()));`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: TB Intakes

## Related VBA Procedures
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form Intakes::cmdLastWeekIntakes_Click [7874-7902]
