# Report Lineage: rptCriminalStatusUpcHrgs

## Trigger Paths
- No trigger path could be inferred from extracted forms/macros/VBA.

## Data Lineage
- RecordSource: `SELECT tblHearingDate.CaseID, tblHearingDate.Hearing_Date, tblHearingDate.HearingType, tblHearingDate.HearingTime, tblHearingDate.Verified, tblHearingDate.HrgResult, tblHearingDate.HrgCal, tblHearingDate.ClientPresent, tblHearingDate.Reminder, tblHearingDate.ReminderCheck, tblHearingDate.HearingID FROM tblHearingDate WHERE (((tblHearingDate.Hearing_Date)>Date()));`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: tblHearingDate

## Related VBA Procedures
- No related VBA procedure could be inferred.
