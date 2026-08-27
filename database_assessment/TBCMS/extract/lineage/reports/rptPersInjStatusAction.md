# Report Lineage: rptPersInjStatusAction

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `SELECT TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, TblActionNeeded.CaseID FROM TblActionNeeded WHERE (((TblActionNeeded.ActionComp)=No));`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: TblActionNeeded

## Related VBA Procedures
- form frmActionNeededAll::cmdActionNeededDone_Click [7129-7159]
- form frmActionNeededAll::Text26_DblClick [7202-7251]
