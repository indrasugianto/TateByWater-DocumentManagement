# Report Lineage: rpt_adj_address_label

## Trigger Paths
- frmPersonalInjury -> cmdPrintAdjLabel -> onClick -> cmdPrintAdjLabel_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT [Personal Injury].Adjuster1, tblCase.CaseID, [Personal Injury].[Adjuster1 Address], [Personal Injury].[Adjuster1 City], [Personal Injury].[Adjuster1 State], [Personal Injury].[Adjuster1 Zip], [Personal Injury].Adjuster1, [Personal Injury].OppPartyInsured, [Personal Injury].InsCo1 FROM tblCase INNER JOIN [Personal Injury] ON tblCase.CaseID = [Personal Injury].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Personal Injury, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16628-16658]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16659-16689]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17559-17696]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17741-17751]
- form frmPersonalInjury::cmdPrintAdjLabel_Click [3802-3817]
- form Intakes::cmdClose_Click [7636-7688]
- form Intakes::cmdCreateOpen_Click [7751-7798]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
