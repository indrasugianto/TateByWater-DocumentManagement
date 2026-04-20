# Report Lineage: rpt_adj_address_label

## Trigger Paths
- frmPersonalInjury -> cmdPrintAdjLabel -> onClick -> cmdPrintAdjLabel_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT [Personal Injury].Adjuster1, tblCase.CaseID, [Personal Injury].[Adjuster1 Address], [Personal Injury].[Adjuster1 City], [Personal Injury].[Adjuster1 State], [Personal Injury].[Adjuster1 Zip], [Personal Injury].Adjuster1, [Personal Injury].OppPartyInsured, [Personal Injury].InsCo1 FROM tblCase INNER JOIN [Personal Injury] ON tblCase.CaseID = [Personal Injury].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Personal Injury, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16447-16477]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16478-16508]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17358-17495]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17540-17550]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660]
- form Intakes::cmdCreateOpen_Click [7723-7770]
- form frmPersonalInjury::cmdPrintAdjLabel_Click [3975-3990]
