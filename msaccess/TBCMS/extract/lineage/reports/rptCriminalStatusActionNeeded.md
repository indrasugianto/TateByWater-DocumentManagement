# Report Lineage: rptCriminalStatusActionNeeded

## Trigger Paths
- No trigger path could be inferred from extracted forms/VBA.

## Data Lineage
- RecordSource: `SELECT TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, TblActionNeeded.CaseID, TblActionNeeded.DateComp, TblActionNeeded.DateComp1 FROM TblActionNeeded WHERE (((TblActionNeeded.ActionComp)=No));`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: TblActionNeeded

## Related VBA Procedures
- form frmActionNeededAll::cmdActionNeededDone_Click [6994-7023]
- form frmActionNeededAll::DateComp1_DblClick [7066-7116]
