# Report Lineage: rpt_adj_address_label

## Trigger Paths
- frmPersonalInjury -> cmdPrintAdjLabel -> onClick -> cmdPrintAdjLabel_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT [Personal Injury].Adjuster1, tblCase.CaseID, [Personal Injury].[Adjuster1 Address], [Personal Injury].[Adjuster1 City], [Personal Injury].[Adjuster1 State], [Personal Injury].[Adjuster1 Zip], [Personal Injury].Adjuster1, [Personal Injury].OppPartyInsured, [Personal Injury].InsCo1 FROM tblCase INNER JOIN [Personal Injury] ON tblCase.CaseID = [Personal Injury].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Personal Injury, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmPersonalInjury::cmdPrintAdjLabel_Click [3975-3990]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
