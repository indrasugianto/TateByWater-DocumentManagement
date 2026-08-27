# Report Lineage: rpt_Main_Closing

## Trigger Paths
- frmClientLedger -> Command353 -> onClick -> Command353_Click (high confidence)
- frmClientLedger -> cmdClosingSheet -> onClick -> cmdClosingSheet_Click (high confidence)
- zClient Ledger OLD -> Command353 -> onClick -> Command353_Click (high confidence)
- frmToBeClosed -> Command353 -> onClick -> Command353_Click (high confidence)
- frmHome -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)
- frmHomeAdmin -> cmdDeleteAll -> onClick -> cmdDeleteAll_Click (high confidence)

## Data Lineage
- RecordSource: `tblCase`
- RecordSourceType: `table-or-unknown`
- Involved Queries: (none)
- Terminal Tables: tblCase

## Related VBA Procedures
- form frmClientLedger::cmdClientReviewEmail_Click [18432-18462] [createObject, runSql]
- form frmClientLedger::cmdClientReviewEmailESP_Click [18463-18493] [createObject, runSql]
- form frmClientLedger::cmdClosingSheet_Click [18529-18540]
- form frmClientLedger::cmbFileNumbers_AfterUpdate [19363-19500] [runSql]
- form frmClientLedger::CaseOpenDate_AfterUpdate [19545-19555]
- form frmClientLedger::Command353_Click [19564-19574]
- form zClient Ledger OLD::Command353_Click [7734-7740]
- form zClient Ledger OLD::cmdSubmitToFamilyLaw_Click [8066-8096] [runSql, setWarnings]
- form zClient Ledger OLD::CaseOpenDate_AfterUpdate [8122-8126]
- form frmToBeClosed::Command353_Click [6856-6863]
- form frmHome::cmdDeleteAll_Click [1868-2005] [runSql]
- form Intakes::cmdClose_Click [7608-7660] [runSql]
- form Intakes::cmdCreateOpen_Click [7723-7770] [runSql]
- form frmHomeAdmin::cmdDeleteAll_Click [6313-6450] [runSql]
- module modGaz::fncGetTABalanceWithCaseID [167-179]
- module modGaz::get_remaining_AdvancedChargesBalance [351-398]
- module modGaz::fncGetMatterARBalanceWithCaseID [406-418]
- module Module1::GetRetainer [20-23]
