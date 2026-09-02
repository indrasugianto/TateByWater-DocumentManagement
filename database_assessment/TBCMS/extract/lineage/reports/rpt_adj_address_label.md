# Report Lineage: rpt_adj_address_label

## Trigger Paths
- frmPersonalInjury -> cmdPrintAdjLabel -> onClick -> cmdPrintAdjLabel_Click (high confidence)

## Data Lineage
- RecordSource: `SELECT [Personal Injury].Adjuster1, tblCase.CaseID, [Personal Injury].[Adjuster1 Address], [Personal Injury].[Adjuster1 City], [Personal Injury].[Adjuster1 State], [Personal Injury].[Adjuster1 Zip], [Personal Injury].Adjuster1, [Personal Injury].OppPartyInsured, [Personal Injury].InsCo1 FROM tblCase INNER JOIN [Personal Injury] ON tblCase.CaseID = [Personal Injury].CaseID;`
- RecordSourceType: `inline-sql`
- Involved Queries: (none)
- Terminal Tables: Personal Injury, tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [16616-16646] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [16647-16677] [createObject, runSql]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [17547-17684] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [17729-17739]
- form frmPersonalInjury::cmdPrintAdjLabel_Click [3802-3817]
- form Intakes::cmdClose_Click [7636-7688] [runSql]
- form Intakes::cmdCreateOpen_Click [7751-7798] [runSql]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8087-8117] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8143-8147]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module DocumentManagement::MoveDocumentByCaseStatus [1373-1585] [createObject, fileSystem]
- module DocumentManagement::GetIntakeDocumentFileName [1682-1758]
- module DocumentManagement::Phase5_E2E_HappyPathTest [1759-2092] [createObject, fileSystem]
- module Module1::GetRetainer [20-23]
