# AI Migration Guide — TB CMS.SQL

> **Audience:** an AI agent (or engineer) producing a plan to rewrite this Microsoft Access application as a web application (React + .NET/Node API + SQL Server). This document explains how the `extract/` folder is organized, how to pull the facts you need out of it, and gives a derived object / data-source map so the relationships between objects are visible up front.

Everything here was produced **statically** — every form/report was opened in Design View only and all code was read via `SaveAsText`. No application code, queries, or macros were executed, so this reflects *structure and intent*, not runtime behavior.

## 1. Snapshot

- **Source database:** `TB CMS.SQL.accdb`
- **Extracted (UTC):** 2026-09-02T15:15:31.985431+00:00
- **Run success:** False

| Object | Count |
|---|---|
| Tables | 93 |
| Relationships (declared) | 0 |
| Queries | 214 |
| Forms | 94 |
| Reports | 99 |
| Macros | 0 |
| VBA procedures | 217 |
| Linked tables | 88 |
| Data-macro tables | 0 |

> ✅ **Complete extraction** — `run_summary.json.completeness` lists are all empty; no objects were skipped. (Still verify `run_summary.json` yourself before treating this as exhaustive.)

## 2. How the extract is structured

Start with `app_manifest.json` — it is the cross-reference hub that links every object. The other files hold the detail the manifest summarizes.

| Path | Contents | Why it matters |
|---|---|---|
| `app_manifest.json` | **Start here.** Object inventory, per-object dependencies, `featureCandidates`, `migrationHints`, `coverage`. | Top-level map of the whole app. |
| `run_summary.json` | Per-stage timing/success + `completeness` (what was skipped). | Trust this before assuming coverage. |
| `schema.json` | Tables → columns (type/size/nullable/default/validation/caption/format/inputMask), PKs, indexes, table-level lookups, calculated/attachment/multi-valued flags, linked-table connect strings. | The data model → SQL Server schema. |
| `tables/<name>.md` + `README.md` | One readable definition per table (columns, PK, indexes, lookups, relationships) rendered from `schema.json`; `README.md` indexes them. | Per-table reference for design work. |
| `relationships.json` | Declared FKs with enforce/cascade flags. | FKs — **may be empty**; see §3. |
| `db_properties.json` | StartUpForm, AppTitle, AccessVersion, security flags. | App entry point + environment. |
| `queries/index.json` + `*.sql` | Query manifest (type, parameters, `referencedTables`) + raw SQL. | Business queries → API endpoints / views. |
| `forms/<name>.json` | Structured form: recordSource, controls (source/format/events), sections, subforms (`resolvedSubform`), `vbaCodeBehind`. | UI screens → React pages/components. |
| `reports/<name>.json` | Structured report: dataSource, group levels, sections, controls, subreports (`resolvedReport`), `vbaCodeBehind`. | Reports → web/PDF reports. |
| `reports/screenshots/*.png` | Design-view image per report (if captured). | Visual reference for layout. |
| `vba/index.json` | Every VBA procedure with kind/line range, `calls`, `usesSql`, `usesDoCmd`, and `redFlags`. | Where the logic lives + risk scan. |
| `vba/forms/ reports/ modules/ classes/` | Raw `SaveAsText` VBA exports. | Full source to port logic from. |
| `macros/<name>.txt` + `index.json` + `ANALYSIS.md` | UI macros parsed into actions with labeled args + per-macro `migrationFlags`; `ANALYSIS.md` is the readable summary. | Navigation/automation logic. |
| `lineage/reports/<name>.json` + `.md` | Trigger paths (which form/control/macro opens a report) + data lineage. | How objects invoke each other. |
| `docs/schema_report.md` | Human-readable schema + inventory dump. | Quick browsing. |

## 3. How to answer the key questions

**What data does a report/form use?** Open `reports/<name>.json` (`dataSource`) or `forms/<name>.json` (`recordSource`). If that value is a query, resolve it via `queries/index.json` → `referencedTables`. Control-level bindings live on each control's `controlSource` / `rowSource`. Subreports/subforms are embedded inline as `resolvedReport` / `resolvedSubform`. See the derived map in §4.

**What are the foreign keys / relationships?** Read `relationships.json` first. If it is empty or thin (common in older `.mdb`s where integrity was never declared), **infer** FKs from: (a) JOIN clauses in `queries/*.sql`, (b) table-level `lookup` definitions in `schema.json` (a combo/list box bound to another table is an implicit FK), and (c) column-name conventions. State these as *inferred* in the plan.

**Where is the business logic?** Three places: (1) `vba/index.json` — procedures with `redFlags`, `usesSql`, `usesDoCmd`; read the full source under `vba/`. (2) `macros/ANALYSIS.md` — flagged macros (`mutatesData`, `businessLogic`, `callsVBA`). (3) `data_macros/` — table-level triggers (`.accdb` only). Logic must move to the API/service layer.

**How do objects invoke each other?** `lineage/reports/<name>.json` gives the trigger path (form → control → event/macro → report). Macro `OpenForm`/`OpenReport` actions in `macros/index.json` show navigation flow. Form/control events (`onClick`, `onCurrent`, …) in `forms/<name>.json` point at the VBA that wires screens together.

**What is risky / needs a manual decision?** `app_manifest.json` → `migrationHints.highRiskAreas`, `complexFields`, `dataMacroTables`, `vbaRedFlagCounts`, and `linkedTableCount`. Summarized in §5.

## 4. Object & data-source map

### Declared relationships

**None declared in the database.** Foreign keys must be *inferred* from query joins, table lookups, and naming (see §3) and added explicitly in the target schema.

### Forms → data source
| Form | Record source | Controls | Events |
|---|---|---|---|
| frmClientLedger | SELECT vwfrmClientLedger.CaseID, vwfrmClientLedger.Last_Name, vwfrmClientLedger.First_Name, vwfrmClientLedger.CaseOpenDate, vwfrmClientLedger.Closed, vwfrmClientLedger.Clsdate, vwfrmClientLedger.Extended_Ledger, vwfrmClientLedger.Case_Letter, vwfrmClientLedger.yr, vwfrmClientLedger.Number_, vwfrmClientLedger.Orig_Atty, vwfrmClientLedger.Address, vwfrmClientLedger.CourtCaseNo, vwfrmClientLedger.City, vwfrmClientLedger.FamilyLaw, vwfrmClientLedger.State, vwfrmClientLedger.Zip, vwfrmClientLedger.Country, vwfrmClientLedger.HmPhone, vwfrmClientLedger.Action, vwfrmClientLedger.OtherPhone, vwfrmClientLedger.Fax, vwfrmClientLedger.WkPhone, vwfrmClientLedger.Comments, vwfrmClientLedger.Email, vwfrmClientLedger.Referral, vwfrmClientLedger.[Individual Referrer], vwfrmClientLedger.Retainer, vwfrmClientLedger.Matter_type, vwfrmClientLedger.SOL, vwfrmClientLedger.Court, vwfrmClientLedger.CType, vwfrmClientLedger.POfc, vwfrmClientLedger.ComplainingWitness, vwfrmClientLedger.DOB, vwfrmClientLedger.WkAddress, vwfrmClientLedger.WkCity, vwfrmClientLedger.WkState, vwfrmClientLedger.WkZip, vwfrmClientLedger.[Pro Bono], vwfrmClientLedger.HandlingAtty_Case, vwfrmClientLedger.Action_Needed_on_Payment, vwfrmClientLedger.SSN, vwfrmClientLedger.[Employer Name], vwfrmClientLedger.[Last Updated Contact Info], vwfrmClientLedger.Ocounsel, vwfrmClientLedger.Firm, vwfrmClientLedger.[OC Address], vwfrmClientLedger.[OC City], vwfrmClientLedger.[OC State], vwfrmClientLedger.[OC Zip], vwfrmClientLedger.[OC Phone], vwfrmClientLedger.[OC Email], vwfrmClientLedger.[OC Fax], vwfrmClientLedger.[Pro Bono PM], vwfrmClientLedger.[Pro Bono JRT], vwfrmClientLedger.ContingencyFee, vwfrmClientLedger.AuthorityToTalkTo, vwfrmClientLedger.Hourly, vwfrmClientLedger.Contingency, vwfrmClientLedger.Hybrid, vwfrmClientLedger.[Family-Law], vwfrmClientLedger.Fixed, vwfrmClientLedger.Scan, vwfrmClientLedger.[Scan Location], vwfrmClientLedger.ScanNotAvail, vwfrmClientLedger.ParaLegal, vwfrmClientLedger.Spanish, vwfrmClientLedger.Offdate, vwfrmClientLedger.CostHold, vwfrmClientLedger.CltNarrative, vwfrmClientLedger.ARTrustZero, vwfrmClientLedger.F73, vwfrmClientLedger.F74, vwfrmClientLedger.F75, vwfrmClientLedger.F76, vwfrmClientLedger.PhName1, vwfrmClientLedger.PhName2, vwfrmClientLedger.LengthRes, vwfrmClientLedger.LengthEmp, vwfrmClientLedger.LegalStatus, vwfrmClientLedger.CurrentBond, vwfrmClientLedger.CrRecord, vwfrmClientLedger.TrustChronMemo, vwfrmClientLedger.Executor, vwfrmClientLedger.RetainerReimb, vwfrmClientLedger.RetReimbAmount, vwfrmClientLedger.Reviewable, vwfrmClientLedger.ReviewReq, vwfrmClientLedger.ReviewReceivedDate, vwfrmClientLedger.ReviewReceived, vwfrmClientLedger.Testimonial, vwfrmClientLedger.ReviewFollowUp, vwfrmClientLedger.Stars, vwfrmClientLedger.[Review Source], vwfrmClientLedger.[Review Date], vwfrmClientLedger.Title, vwfrmClientLedger.OPartyLast, vwfrmClientLedger.OPartyFirst, vwfrmClientLedger.OPartyDOB, vwfrmClientLedger.SSMA_TimeStamp, vwfrmClientLedger.FileNo, vwfrmClientLedger.PartnerRate, vwfrmClientLedger.AssocRate FROM vwfrmClientLedger WHERE (((1)=0));  | 341 | 4 |
| frm_advanced_payments | qry_advanced_payments | 38 | 0 |
| frm_uncashed_trust_checks | qry_uncashed_trust_checks | 30 | 0 |
| frmActionNeededAll3 | qryActionNeededAll | 30 | 0 |
| frmBankruptcy | SELECT Bankruptcy.BankruptcyID, Bankruptcy.CaseID, Bankruptcy.Chapter, Bankruptcy.[Case Filed], Bankruptcy.[Deadline for Filing Sched], Bankruptcy.[Document Date for Trustee], Bankruptcy.Trustee, Bankruptcy.POCDeadline, Bankruptcy.GovtPOC, Bankruptcy.[Deadline to Object], Bankruptcy.BJudge, Bankruptcy.OriginalScheduleDeadline, Bankruptcy.PrevBank, Bankruptcy.PrevDate, Bankruptcy.PrevCaseNumber, Bankruptcy.PrevLocation, Bankruptcy.TrusteeAddress, Bankruptcy.TrusteeCity, Bankruptcy.TrusteeZip, Bankruptcy.TrusteeState, Bankruptcy.TrusteeDocuments, Bankruptcy.ForeTrustee, Bankruptcy.ForeAddress, Bankruptcy.ForeCity, Bankruptcy.ForeState, Bankruptcy.ForeZIP, Bankruptcy.ForePhone, Bankruptcy.ForeFax, Bankruptcy.ForeSaleDate, Bankruptcy.ForeTime, Bankruptcy.ForeFileNumber, Bankruptcy.TrusteePhone, Bankruptcy.TrusteeFax, Bankruptcy.TrusteeEmail, Bankruptcy.SSMA_TimeStamp FROM Bankruptcy;  | 68 | 0 |
| frm_trust_summary | qry_trustStatements | 30 | 1 |
| frm_invoices_summary | SELECT vw_frm_invoices_summary.CaseID, vw_frm_invoices_summary.Name, vw_frm_invoices_summary.First_Name, vw_frm_invoices_summary.Last_Name, vw_frm_invoices_summary.Retainer, vw_frm_invoices_summary.SumOfCharge, vw_frm_invoices_summary.SumOfPayment, vw_frm_invoices_summary.SumOfBalance, vw_frm_invoices_summary.BalanceCalculated, vw_frm_invoices_summary.BalRetCalculated, vw_frm_invoices_summary.FileNumber, vw_frm_invoices_summary.[Balance Due Date], vw_frm_invoices_summary.Orig_Atty, vw_frm_invoices_summary.HandlingAtty_Case, vw_frm_invoices_summary.CodeVal, vw_frm_invoices_summary.Executor, vw_frm_invoices_summary.LastOfInvSent, * FROM vw_frm_invoices_summary ORDER BY vw_frm_invoices_summary.[Balance Due Date] DESC;  | 58 | 1 |
| frm_Billing_Tracker | qryBillingTracker | 19 | 0 |
| frmDispositions | qryDispos | 52 | 0 |
| frm_Billing_Tracker2 | qryBillingTracker2 | 36 | 0 |
| frmTimeTableDetailMerge | SELECT tblTimeTableDetail.Time_ID, tblTimeTableDetail.Tdate, tblTimeTableDetail.Description, tblTimeTableDetail.Tatty, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, tblTimeTableDetail.Bill_ID FROM tblTimeTableDetail ORDER BY tblTimeTableDetail.Time_ID;  | 22 | 1 |
| frmChild | SELECT tblChild.Child_ID, tblChild.FamilyLaw_ID, tblChild.ChildName, tblChild.DOB_child FROM tblChild;  | 7 | 0 |
| frmActionNeeded | SELECT TblActionNeeded.ActionNeededID, TblActionNeeded.CaseID, TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, TblActionNeeded.SSMA_TimeStamp, TblActionNeeded.DateComp, TblActionNeeded.DateComp1, TblActionNeeded.ActPerson, TblActionNeeded.StartDate FROM TblActionNeeded;  | 9 | 1 |
| frmTakeOffReconciliation | qryTakeOff | 121 | 1 |
| frmAttyFeeGeneration | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM tblTakeOffMonth ORDER BY tblTakeOffMonth.TakeOffDate;  | 101 | 0 |
| frmActionNeededAll | qryActionNeededAll | 39 | 0 |
| frmActionNeededAll2 | qryActionNeededAll2 | 27 | 0 |
| frmCalls | tblCalls | 70 | 1 |
| frmLogin | — | 10 | 2 |
| frmClientReviews | SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, tblCase.ReviewFollowUp, tblCase.ReviewReceivedDate, tblCase.ReviewReq, tblCase.Reviewable, tblCase.ReviewReceived, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Last_Name] & ", " & [First_Name] AS Name, tblCase.CaseID, tblCase.ParaLegal, tblCase.Testimonial, tblCase.Stars, tblCase.[Review Source] FROM tblCase WHERE (((tblCase.Reviewable)=True)) ORDER BY tblCase.ReviewReq DESC;  | 37 | 1 |
| frmAddUser | tblUsers | 11 | 1 |
| frmCalendarCheck | qryCalendarCheck | 39 | 0 |
| frmAdminLoginTK | — | 8 | 0 |
| frmPersInjDemand | tblPersInjDemand | 9 | 0 |
| frmHearingDate | SELECT tblHearingDate.HearingID, tblHearingDate.CaseID, tblHearingDate.Hearing_Date, tblHearingDate.HearingType, tblHearingDate.HearingTime, tblHearingDate.HrgResult, tblHearingDate.HrgCal, tblHearingDate.Verified, tblHearingDate.ClientPresent, tblHearingDate.Reminder, tblHearingDate.ReminderCheck, tblHearingDate.SSMA_TimeStamp FROM tblHearingDate ORDER BY tblHearingDate.Hearing_Date;  | 21 | 0 |
| frmApplicationLoad | — | 0 | 1 |
| frmCaseList | — | 11 | 0 |
| frmAttyNotes | SELECT tblNotes.IDNotes, tblNotes.CaseID, tblNotes.NoteDate, tblNotes.NotePerson, tblNotes.NoteDescription, tblNotes.NoteTime, tblNotes.SSMA_TimeStamp FROM tblNotes ORDER BY tblNotes.NoteDate;  | 10 | 0 |
| frmBilling | SELECT Billing.ID, Billing.CaseID, Billing.[Balance Due Date], Billing.[Past Due], Billing.[Long Term Collections], Billing.chkBalanceDue, Billing.[Billing Notes], Billing.WriteOff, Billing.CostHold, Billing.SSMA_TimeStamp FROM Billing;  | 14 | 0 |
| frmBrowse | — | 5 | 1 |
| frmCaseListClosed | qryCaseListClosed | 34 | 1 |
| frmBrowse_BackEnd | — | 5 | 1 |
| frmReceipt | tblReceipts | 33 | 1 |
| frmClientsConflict | tblCase | 17 | 0 |
| frmUsers_Edit | SELECT * FROM tblUsers;  | 12 | 1 |
| frmCallsList | SELECT tblCalls.CFirstName, tblCalls.CLastName, tblCalls.CDate, tblCalls.CallTime, tblCalls.CPracticeArea, tblCalls.CReferral, tblCalls.CAtty, tblCalls.CallMatter, tblCalls.CPhone, tblCalls.ClientType, tblCalls.CallID, tblDropD.CodeVal, Nz([CFirstName])+" "+Nz([CLastName]) AS CallName FROM tblCalls LEFT JOIN tblDropD ON tblCalls.CPracticeArea = tblDropD.Code ORDER BY tblCalls.CDate DESC;  | 39 | 0 |
| frmCaseListOpen subform | SELECT [qryCaseListOpen].[CaseID], [qryCaseListOpen].[CaseOpenDate], [qryCaseListOpen].[ClientName], [qryCaseListOpen].[Case_Letter], [qryCaseListOpen].[yr], [qryCaseListOpen].[Number_], [qryCaseListOpen].[Orig_Atty], [qryCaseListOpen].[Extended_Ledger], [qryCaseListOpen].[Court], [qryCaseListOpen].[Matter_type], [qryCaseListOpen].[FileNumber], [qryCaseListOpen].[Scan Location], [qryCaseListOpen].[HandlingAtty_Case], [qryCaseListOpen].[Closed], [qryCaseListOpen].[CodeVal] FROM qryCaseListOpen;  | 30 | 0 |
| frmCaseListAll | qryCaseListAll | 38 | 1 |
| frmCaseListOpen | qryCaseListOpen | 39 | 1 |
| frmHomeAdmin | — | 12 | 0 |
| zfrmSelectCaseNum | — | 4 | 0 |
| frmConflictChk | — | 8 | 0 |
| frmPersInjuryStatusReport | — | 7 | 0 |
| frmOppPartyConflict | tblCase | 15 | 0 |
| frmFamilyLaw | SELECT [Family Law - Divorce].ID, [Family Law - Divorce].CaseID, [Family Law - Divorce].[C Length at Residence], [Family Law - Divorce].[C Prior Address], [Family Law - Divorce].[C Length at Prior Address], [Family Law - Divorce].[C Length in VA], [Family Law - Divorce].[C Birthplace], [Family Law - Divorce].[C Employer], [Family Law - Divorce].[C Primary Education], [Family Law - Divorce].[C College], [Family Law - Divorce].[C Marriage Number], [Family Law - Divorce].[D Address], [Family Law - Divorce].[D City], [Family Law - Divorce].[D State], [Family Law - Divorce].[D Zip], [Family Law - Divorce].[D Home Phone], [Family Law - Divorce].[D Other Phone], [Family Law - Divorce].[D Email], [Family Law - Divorce].[D DOB], [Family Law - Divorce].[D SSN], [Family Law - Divorce].[D Employer], [Family Law - Divorce].[D Work Address], [Family Law - Divorce].[D Work City], [Family Law - Divorce].[D Work State], [Family Law - Divorce].[D Work Zip], [Family Law - Divorce].[D Work Phone], [Family Law - Divorce].[D Primary Education], [Family Law - Divorce].[D College], [Family Law - Divorce].[D Marriage Number], [Family Law - Divorce].[Date of Marriage], [Family Law - Divorce].[Place of Marriage], [Family Law - Divorce].[Date of Separation], [Family Law - Divorce].[Length of Separation], [Family Law - Divorce].[Wife Maiden Name], [Family Law - Divorce].[Number of Children], [Family Law - Divorce].[Child Custodian], [Family Law - Divorce].[C Title], [Family Law - Divorce].[D Title], [Family Law - Divorce].[Date of PSA], [Family Law - Divorce].[Place of Last Cohabit], [Family Law - Divorce].[Divorce Grounds], [Family Law - Divorce].[FL Court Case No], [Family Law - Divorce].[Complaint Filed Date], [Family Law - Divorce].[Waiver Date], [Family Law - Divorce].[Publish Dates], [Family Law - Divorce].[Publish Return Date], [Family Law - Divorce].[Complaint Serve Date], [Family Law - Divorce].[Complaint Serve Method], [Family Law - Divorce].[NOH Serve Date], [Family Law - Divorce].[NOH Serve Method], [Family Law - Divorce].Witness, [Family Law - Divorce].D_Last_Name, [Family Law - Divorce].D_First_Name, [Family Law - Divorce].D_BirthPlace, [Family Law - Divorce].[Uncontested by Affidavit], [Family Law - Divorce].[Waiver of Service], [Family Law - Divorce].[Service by Publication], [Family Law - Divorce].[Sheriff Service], [Family Law - Divorce].[Divorce with MSA], [Family Law - Divorce].[Divorce without MSA], [Family Law - Divorce].SSMA_TimeStamp FROM [Family Law - Divorce];  | 127 | 1 |
| frmOpenReport | — | 10 | 0 |
| frmCrimStatusReport | — | 11 | 0 |
| frmCtCaseNumbers | SELECT tbl_CtCaseNumbers.CtCaseNoID, tbl_CtCaseNumbers.CaseID, tbl_CtCaseNumbers.Matter_Charge, tbl_CtCaseNumbers.CtNumber, tbl_CtCaseNumbers.District, tbl_CtCaseNumbers.Circuit, tbl_CtCaseNumbers.CodeSection, tbl_CtCaseNumbers.SSMA_TimeStamp FROM tbl_CtCaseNumbers;  | 6 | 0 |
| frmDisposition | SELECT Disposition.DispoID, Disposition.CaseID, Disposition.Disposition, Disposition.Trial, Disposition.Plea, Disposition.[Not Guilty Dismissed], Disposition.[Entire np], Disposition.[PI Settlement Amount], Disposition.Dispo_Date, Disposition.Dispo_Atty, Disposition.DispoJudge, Disposition.DispoOppC, Disposition.[Total Earned Fee], Disposition.SSMA_TimeStamp FROM Disposition;  | 25 | 0 |
| frmHome | — | 38 | 3 |
| frmScanLocation | tblScans | 3 | 0 |
| frmHomeAdminLogin | — | 8 | 0 |
| frmOkAlert | — | 4 | 1 |
| frmIntakesConflicts | TB Intakes | 17 | 0 |
| frmSubProofOfClaims | SELECT ProofOfClaims.IDProofOfClaims, ProofOfClaims.IDBankruptcy, ProofOfClaims.ClaimNr, ProofOfClaims.DateFiled, ProofOfClaims.CreditorName, ProofOfClaims.Secured, ProofOfClaims.Priority, ProofOfClaims.Unsecured, ProofOfClaims.Arrears FROM ProofOfClaims;  | 23 | 0 |
| frmInvoiceSent | SELECT tbl_InvoiceSent.CaseID, tbl_InvoiceSent.InvSent, tbl_InvoiceSent.InvoiceNumber, tbl_InvoiceSent.[TK Sent], tbl_InvoiceSent.TKDate, tbl_InvoiceSent.InvoiceSentID, tbl_InvoiceSent.InvSentNotes, tbl_InvoiceSent.InvBalance, tbl_InvoiceSent.TKNumber, tbl_InvoiceSent.TKBalance, tbl_InvoiceSent.ClientCall, tbl_InvoiceSent.InvoiceSentID FROM tbl_InvoiceSent ORDER BY tbl_InvoiceSent.[TK Sent], tbl_InvoiceSent.TKDate;  | 18 | 1 |
| frmPersonalInjury2 | Personal Injury | 107 | 0 |
| frmMatter | SELECT vwMatterAndAR.MatterID, vwMatterAndAR.CaseID, vwMatterAndAR.Date2, vwMatterAndAR.Pay_Outlay, vwMatterAndAR.Charge, vwMatterAndAR.Payment, vwMatterAndAR.FirmPrepaid, vwMatterAndAR.OrderNr, vwMatterAndAR.InsertPymt, vwMatterAndAR.AdvancedLegal, vwMatterAndAR.SumOfCharge, vwMatterAndAR.SumOfPayment, vwMatterAndAR.Retainer, vwMatterAndAR.Balance, vwMatterAndAR.Creimb FROM vwMatterAndAR ORDER BY vwMatterAndAR.CaseID, vwMatterAndAR.OrderNr;  | 17 | 0 |
| frmPersInjLog | SELECT tblPersInjLog.EventDate, tblPersInjLog.EventDescription, tblPersInjLog.ID, tblPersInjLog.PersInjLogID, tblPersInjLog.LogParalegal FROM tblPersInjLog ORDER BY tblPersInjLog.EventDate DESC;  | 10 | 1 |
| frmPersInjProvider | SELECT tblPersInjProv.PIProviderID, tblPersInjProv.ID, tblPersInjProv.Provider, tblPersInjProv.ReqDate, tblPersInjProv.RcvDate, tblPersInjProv.PBillAmount, tblPersInjProv.Lien, tblPersInjProv.SSMA_TimeStamp, tblPersInjProv.PBillRed FROM tblPersInjProv;  | 21 | 0 |
| frmPersInjLog2 | tblPersInjLog | 6 | 0 |
| frmPersonalInjury | SELECT [Personal Injury].ID, [Personal Injury].CaseID, [Personal Injury].ClaimNo1, [Personal Injury].InsCo1, [Personal Injury].Adjuster1, [Personal Injury].[Adjuster1 Address], [Personal Injury].[Adjuster1 City], [Personal Injury].[Adjuster1 State], [Personal Injury].[Adjuster1 Zip], [Personal Injury].[Adjuster1 Phone], [Personal Injury].[Adjuster1 Fax], [Personal Injury].[Adjuster1 Email], [Personal Injury].ClaimNo2, [Personal Injury].InsCo2, [Personal Injury].Adjuster2, [Personal Injury].[Adjuster2 Address], [Personal Injury].[Adjuster2 City], [Personal Injury].[Adjuster2 State], [Personal Injury].[Adjuster2 Zip], [Personal Injury].[Adjuster2 Phone], [Personal Injury].[Adjuster2 Fax], [Personal Injury].[Adjuster2 Email], [Personal Injury].[Filing Date], [Personal Injury].Medicare, [Personal Injury].[Med Pay], [Personal Injury].ERISA, [Personal Injury].Litigation, [Personal Injury].[Slip and Fall], [Personal Injury].[Auto Accident], [Personal Injury].[Medical Lien], [Personal Injury].Assignment, [Personal Injury].[Med Mal], [Personal Injury].DOI, [Personal Injury].HealthIns, [Personal Injury].PolicyNo, [Personal Injury].GroupNo, [Personal Injury].csettleper, [Personal Injury].csettlelit, [Personal Injury].location, [Personal Injury].Medicaid, [Personal Injury].OppPartyInsured, [Personal Injury].Demand, [Personal Injury].BriefDescription, [Personal Injury].PIState, [Personal Injury].AutoCarrier, [Personal Injury].AutoPolicyNo, [Personal Injury].UnderinsLimits, [Personal Injury].MaxMed, [Personal Injury].PISOL, [Personal Injury].PolicyNo1, [Personal Injury].AdjusterExt, [Personal Injury].CompltServed, [Personal Injury].ServedDate, [Personal Injury].OtherDriver, [Personal Injury].SSMA_TimeStamp, [Personal Injury].PIStatus, [Personal Injury].LiabilityLimit FROM [Personal Injury];  | 109 | 0 |
| frmScansubform | SELECT tblScans.ScansID, tblScans.CaseID, tblScans.ScanLocation, tblScans.TypeofScan, tblScans.SSMA_TimeStamp FROM tblScans;  | 2 | 0 |
| frmSourceAnalytics | qryCaseSourcesRPT1 | 45 | 0 |
| frmSubCH13Plans | SELECT CH13Plans.IDCH13Plans, CH13Plans.IDBankruptcy, CH13Plans.PlanNr, CH13Plans.DateFiled, CH13Plans.ConfirmDate, CH13Plans.Notes, CH13Plans.Confirmed, CH13Plans.Objected, CH13Plans.SSMA_TimeStamp FROM CH13Plans;  | 12 | 0 |
| frmSubPrevBankrupt | SELECT tblPrevBank.IDPrevBank, tblPrevBank.IDBankruptcy, tblPrevBank.PrevDate, tblPrevBank.PrevCaseNumber, tblPrevBank.PrevLocation, tblPrevBank.PChapter FROM tblPrevBank;  | 10 | 0 |
| frmTakeOff | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM tblTakeOffMonth ORDER BY tblTakeOffMonth.TakeOffDate DESC;  | 154 | 0 |
| frmTakeOff2 | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM tblTakeOffMonth ORDER BY tblTakeOffMonth.TakeOffDate DESC;  | 9 | 0 |
| frmTakeOffSteps | — | 15 | 0 |
| frmTakeOffSubForm | SELECT vwfrmTakeOffSubForm.FileNumber, vwfrmTakeOffSubForm.Name, vwfrmTakeOffSubForm.CaseID, vwfrmTakeOffSubForm.Last_Name, vwfrmTakeOffSubForm.First_Name, vwfrmTakeOffSubForm.CaseOpenDate, vwfrmTakeOffSubForm.Closed, vwfrmTakeOffSubForm.Clsdate, vwfrmTakeOffSubForm.Extended_Ledger, vwfrmTakeOffSubForm.Case_Letter, vwfrmTakeOffSubForm.yr, vwfrmTakeOffSubForm.Number_, vwfrmTakeOffSubForm.Orig_Atty, vwfrmTakeOffSubForm.Address, vwfrmTakeOffSubForm.CourtCaseNo, vwfrmTakeOffSubForm.City, vwfrmTakeOffSubForm.FamilyLaw, vwfrmTakeOffSubForm.State, vwfrmTakeOffSubForm.Zip, vwfrmTakeOffSubForm.Country, vwfrmTakeOffSubForm.HmPhone, vwfrmTakeOffSubForm.Action, vwfrmTakeOffSubForm.OtherPhone, vwfrmTakeOffSubForm.Fax, vwfrmTakeOffSubForm.WkPhone, vwfrmTakeOffSubForm.Comments, vwfrmTakeOffSubForm.Email, vwfrmTakeOffSubForm.Referral, vwfrmTakeOffSubForm.[Individual Referrer], vwfrmTakeOffSubForm.Retainer, vwfrmTakeOffSubForm.Matter_type, vwfrmTakeOffSubForm.SOL, vwfrmTakeOffSubForm.Court, vwfrmTakeOffSubForm.CType, vwfrmTakeOffSubForm.POfc, vwfrmTakeOffSubForm.ComplainingWitness, vwfrmTakeOffSubForm.DOB, vwfrmTakeOffSubForm.WkAddress, vwfrmTakeOffSubForm.WkCity, vwfrmTakeOffSubForm.WkState, vwfrmTakeOffSubForm.WkZip, vwfrmTakeOffSubForm.[Pro Bono], vwfrmTakeOffSubForm.HandlingAtty_Case, vwfrmTakeOffSubForm.Action_Needed_on_Payment, vwfrmTakeOffSubForm.SSN, vwfrmTakeOffSubForm.[Employer Name], vwfrmTakeOffSubForm.[Last Updated Contact Info], vwfrmTakeOffSubForm.Ocounsel, vwfrmTakeOffSubForm.Firm, vwfrmTakeOffSubForm.[OC Address], vwfrmTakeOffSubForm.[OC City], vwfrmTakeOffSubForm.[OC State], vwfrmTakeOffSubForm.[OC Zip], vwfrmTakeOffSubForm.[OC Phone], vwfrmTakeOffSubForm.[OC Email], vwfrmTakeOffSubForm.[OC Fax], vwfrmTakeOffSubForm.[Pro Bono PM], vwfrmTakeOffSubForm.[Pro Bono JRT], vwfrmTakeOffSubForm.ContingencyFee, vwfrmTakeOffSubForm.AuthorityToTalkTo, vwfrmTakeOffSubForm.Hourly, vwfrmTakeOffSubForm.Contingency, vwfrmTakeOffSubForm.Hybrid, vwfrmTakeOffSubForm.[Family-Law], vwfrmTakeOffSubForm.Fixed, vwfrmTakeOffSubForm.Scan, vwfrmTakeOffSubForm.[Scan Location], vwfrmTakeOffSubForm.ScanNotAvail, vwfrmTakeOffSubForm.ParaLegal, vwfrmTakeOffSubForm.Spanish, vwfrmTakeOffSubForm.Offdate, vwfrmTakeOffSubForm.CostHold, vwfrmTakeOffSubForm.CltNarrative, vwfrmTakeOffSubForm.ARTrustZero, vwfrmTakeOffSubForm.F73, vwfrmTakeOffSubForm.F74, vwfrmTakeOffSubForm.F75, vwfrmTakeOffSubForm.F76, vwfrmTakeOffSubForm.PhName1, vwfrmTakeOffSubForm.PhName2, vwfrmTakeOffSubForm.LengthRes, vwfrmTakeOffSubForm.LengthEmp, vwfrmTakeOffSubForm.LegalStatus, vwfrmTakeOffSubForm.CurrentBond, vwfrmTakeOffSubForm.CrRecord, vwfrmTakeOffSubForm.TrustChronMemo, vwfrmTakeOffSubForm.Executor, vwfrmTakeOffSubForm.RetainerReimb, vwfrmTakeOffSubForm.RetReimbAmount, vwfrmTakeOffSubForm.Reviewable, vwfrmTakeOffSubForm.ReviewReq, vwfrmTakeOffSubForm.ReviewReceivedDate, vwfrmTakeOffSubForm.ReviewReceived, vwfrmTakeOffSubForm.Testimonial, vwfrmTakeOffSubForm.ReviewFollowUp, vwfrmTakeOffSubForm.Stars, vwfrmTakeOffSubForm.[Review Source], vwfrmTakeOffSubForm.[Review Date], vwfrmTakeOffSubForm.Title, vwfrmTakeOffSubForm.OPartyLast, vwfrmTakeOffSubForm.OPartyFirst, vwfrmTakeOffSubForm.OPartyDOB, vwfrmTakeOffSubForm.TakeOffID, vwfrmTakeOffSubForm.TakeOffMonthID, vwfrmTakeOffSubForm.AvailBalance, vwfrmTakeOffSubForm.TotalUnCashedChks, vwfrmTakeOffSubForm.TotalUnclearedDeps, vwfrmTakeOffSubForm.TotalAdvancedAR, vwfrmTakeOffSubForm.EarlyEarned, vwfrmTakeOffSubForm.TOEarned, vwfrmTakeOffSubForm.CostReimb, vwfrmTakeOffSubForm.CBHRev, vwfrmTakeOffSubForm.MKRev, vwfrmTakeOffSubForm.CBHCom, vwfrmTakeOffSubForm.MTRev, vwfrmTakeOffSubForm.MTCom, vwfrmTakeOffSubForm.KBCom, vwfrmTakeOffSubForm.MKCom, vwfrmTakeOffSubForm.TOEarnedTr, vwfrmTakeOffSubForm.CostReimbTr, vwfrmTakeOffSubForm.InsertedTrust, vwfrmTakeOffSubForm.TotalHourlyOuts, vwfrmTakeOffSubForm.OpenTK, vwfrmTakeOffSubForm.AdvCostBal, vwfrmTakeOffSubForm.AdvFeeBal, vwfrmTakeOffSubForm.CostHoldBal, vwfrmTakeOffSubForm.BRRev, vwfrmTakeOffSubForm.BRCom, vwfrmTakeOffSubForm.RLFCom, vwfrmTakeOffSubForm.AdvEarned, vwfrmTakeOffSubForm.RemEarned, vwfrmTakeOffSubForm.SumOfCBHRev, vwfrmTakeOffSubForm.SumOfMKRev, vwfrmTakeOffSubForm.SumOfCBHCom, vwfrmTakeOffSubForm.SumOfMTRev, vwfrmTakeOffSubForm.SumOfMTCom, vwfrmTakeOffSubForm.SumOfKBCom, vwfrmTakeOffSubForm.SumOfMKCom, vwfrmTakeOffSubForm.SumOfRLFCom, vwfrmTakeOffSubForm.SumOfEarlyEarned, vwfrmTakeOffSubForm.SumOfTOEarned, vwfrmTakeOffSubForm.SumOfTOEarlyAndEarned, vwfrmTakeOffSubForm.SumOfCostReimb, vwfrmTakeOffSubForm.TOAttBilled FROM vwfrmTakeOffSubForm ORDER BY vwfrmTakeOffSubForm.Name;  | 84 | 0 |
| frmTakeOffSubForm_OLD | qryTakeOffStep2 | 77 | 0 |
| frmTakeOffSubForm2 | qryTakeOffStep2 | 63 | 0 |
| frmTakeOffSubForm3 | qryTakeOffStep2 | 75 | 0 |
| frmTakeOffTest | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM tblTakeOffMonth ORDER BY tblTakeOffMonth.TakeOffDate DESC;  | 9 | 0 |
| frmTakeOffTotalFeesCosts | tblTakeOffMonth | 39 | 0 |
| frmTimeKeepingClosed | qryTimeKeepingClosed | 53 | 0 |
| frmTimeKeepingOpen | qryTimeKeepingOpen | 37 | 0 |
| frmTimeTableDetail | SELECT vwTimeTableDetail.Time_ID, vwTimeTableDetail.Tdate, vwTimeTableDetail.Description, vwTimeTableDetail.Tatty, vwTimeTableDetail.Rate, vwTimeTableDetail.Time_, vwTimeTableDetail.Bill_ID, vwTimeTableDetail.Amount FROM vwTimeTableDetail ORDER BY vwTimeTableDetail.Tdate;  | 31 | 1 |
| frmTKClose | qryTKClose1 | 71 | 0 |
| frmToBeClosed | qryToBeClosed | 32 | 0 |
| frmToBeScanned | qryToBeScanned | 31 | 0 |
| frmTrustAccount | SELECT vwTrustAccountTable.TrustAccountID, vwTrustAccountTable.CaseID, vwTrustAccountTable.TDate, vwTrustAccountTable.TMatter, vwTrustAccountTable.Debit, vwTrustAccountTable.Credit, vwTrustAccountTable.CheckCashed, vwTrustAccountTable.CheckNumber, vwTrustAccountTable.DepCleared, vwTrustAccountTable.Reconciled, vwTrustAccountTable.OrderNr, vwTrustAccountTable.CostReimb, vwTrustAccountTable.AdvFee, vwTrustAccountTable.SumOfDebit, vwTrustAccountTable.SumOfCredit, vwTrustAccountTable.Balance FROM vwTrustAccountTable ORDER BY vwTrustAccountTable.CaseID, vwTrustAccountTable.OrderNr;  | 17 | 1 |
| zfrmSelectCaseNum_Discount | — | 4 | 0 |
| frmTRUSTENTRIESCHRON | qryTrustEntriesChron | 68 | 1 |
| frmUpcoming Hearings | qryUpcomingHearings | 53 | 0 |
| frmUsers | SELECT * FROM tblUsers;  | 10 | 1 |
| frmYearWiseCaseList | SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, TblCase.Extended_Ledger, TblCase.Court, TblCase.Matter_type, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.[Scan Location], TblCase.HandlingAtty_Case, TblCase.Closed, tblDropD.CodeVal FROM TblCase LEFT JOIN tblDropD ON TblCase.Case_Letter = tblDropD.Code ORDER BY TblCase.Number_;  | 41 | 1 |
| frmYesNoAlert | — | 5 | 1 |
| Intakes | TB Intakes | 59 | 0 |
| Time Keeping | qryTimeKeeping | 113 | 3 |
| zClient Ledger OLD | tblCase | 209 | 1 |
| zfrmFamilyLaw OLD | qryFamilyLaw | 164 | 1 |
| zfrmPersInjSOL | qrySOL | 27 | 0 |
| zfrmPersonalDetailsFamilyLaw | tblCase | 30 | 1 |

### Reports → data source
| Report | Data source | Subreports | Code-behind |
|---|---|---|---|
| rpt_TKTotalAdvance | qryInvoiceAttachRPT1 | — | yes |
| rpt_Matter_Closing | SELECT vw_rpt_Matter_Closing.CaseID, vw_rpt_Matter_Closing.MatterID, vw_rpt_Matter_Closing.Date2, vw_rpt_Matter_Closing.Pay_Outlay, vw_rpt_Matter_Closing.Charge, vw_rpt_Matter_Closing.Payment, vw_rpt_Matter_Closing.Balance, vw_rpt_Matter_Closing.RunningDebit, vw_rpt_Matter_Closing.RunningCredit, vw_rpt_Matter_Closing.RunningBalance, vw_rpt_Matter_Closing.Retainer, vw_rpt_Matter_Closing.RetBal, vw_rpt_Matter_Closing.OrderNr FROM vw_rpt_Matter_Closing ORDER BY vw_rpt_Matter_Closing.MatterID;  | — | yes |
| rptCriminalStatusNotesLog | SELECT tblNotes.CaseID, tblNotes.NoteDate, tblNotes.NotePerson, tblNotes.NoteDescription, tblNotes.NoteTime, tblNotes.CaseID, tblNotes.IDNotes FROM tblNotes;  | — | yes |
| rpt_Trust_Closing | qryStmtTrustRPT1 | — | yes |
| rpt_Billing_Closing | SELECT tblCase.CaseID, Disposition.[Total Earned Fee] AS Expr1 FROM tblCase INNER JOIN Disposition ON tblCase.CaseID = Disposition.CaseID;  | — | yes |
| Accounts Receivable | qry_invoices_summaryRPT | — | yes |
| rpt_Main_Closing | tblCase | Report.rpt_Disposition_Closing, Report.rpt_Matter_Closing, Report.rpt_Trust_Closing, Report.rpt_Billing_Closing, Report.rpt_CaseNumber_Closing | yes |
| rpt_Disposition_Closing | qry_Disposition_ClosingSheet | — | yes |
| rpt_CaseNumber_Closing | tbl_CtCaseNumbers | — | yes |
| Case Sources and Revenue | qryCaseSourcesRPT1 | — | yes |
| rpt_Comprehensive_InvoiceStmtS | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, [TB Time Keeping].AdvBalanceatClose, [TB Time Keeping].TrustatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceTKEx3Costs | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, tblCase.Executor, [TB Time Keeping].AdvBalanceatClose FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| Copy Of Client Closing Sheet | qryClosing RPT1 | — | yes |
| rpt_Compr_InvoiceADVCur | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprARCur, Report.rptInvoiceComprPymtsARCur, Report.rptInvoiceComprTrustCur | yes |
| Client Closing Sheet | qryClosing RPT1 | — | yes |
| rptInvoiceComprehensiveTrust2 | qryInvoiceComprehensiveTrustCredit4 | — | yes |
| Client_Trust_Accounts_for_PreTake_Off | qryTakeOff | — | yes |
| Client_Trust_Accounts_for_Take_Off | qryAttyTrustAcctsTOff | — | yes |
| rpt_Reconciliation sub | SELECT qryTakeOffStep2.FileNumber, qryTakeOffStep2.Name, qryTakeOffStep2.AvailBalance, qryTakeOffStep2.TakeOffMonthID FROM qryTakeOffStep2;  | — | yes |
| rpt_Comprehensive_Invoice | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].[BilL Closed Date] FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rpt_TimeDetail_Comprehensive, Report.rptInvoiceComprehensiveAR, Report.rptInvoiceComprehensiveTrust | yes |
| New Invoice | qry_current_invoice | — | yes |
| Invoice | qryInvoiceRPT1 | — | yes |
| Invoice - No Balance Due | qryInvoiceRPT1 | — | yes |
| Invoice - Past Due | qryInvoiceRPT1 | — | yes |
| Invoice Attach - Hourly | qryInvoiceAttachRPT1 | — | yes |
| rptPISOLList | — | — | yes |
| rpt_Comprehensive_InvoiceTKEx3CostsS | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, tblCase.Executor, [TB Time Keeping].AdvBalanceatClose FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| Invoice Attach - Hourly w Discount | qryInvoiceAttachRPT1 | — | yes |
| rpt_Trust_Chron_35 | qryTrustEntriesChronRPT35 | — | yes |
| rptInvoiceComprehensiveAR2 | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, tblCase.Retainer, tblCase.CaseOpenDate, [Matter and AR].OrderNr, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Bill_ID FROM (tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID) INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID WHERE ((([Matter and AR].Date2)<=[Bill Closed Date]) AND (([Matter and AR].Charge)<>0)) ORDER BY [Matter and AR].Date2;  | — | yes |
| Invoice2 | qryInvoiceRPT1 | — | yes |
| rpt_address_label | tblCase | — | yes |
| rpt_address_labelEx | tblCase | — | yes |
| rptLastTenOpen | SELECT qryCaseListOpen.CaseID, qryCaseListOpen.CaseOpenDate, qryCaseListOpen.ClientName, qryCaseListOpen.Orig_Atty, qryCaseListOpen.Matter_type, qryCaseListOpen.FileNumber, qryCaseListOpen.Retainer, qryCaseListOpen.Number_, qryCaseListOpen.yr, qryCaseListOpen.Case_Letter, tblCase.Referral, qryCaseListOpen.CodeVal FROM qryCaseListOpen INNER JOIN tblCase ON qryCaseListOpen.CaseID = tblCase.CaseID WHERE (((qryCaseListOpen.CaseOpenDate) Between getSTDT() And getENDT()));  | — | yes |
| rpt_adj_address_label | SELECT [Personal Injury].Adjuster1, tblCase.CaseID, [Personal Injury].[Adjuster1 Address], [Personal Injury].[Adjuster1 City], [Personal Injury].[Adjuster1 State], [Personal Injury].[Adjuster1 Zip], [Personal Injury].Adjuster1, [Personal Injury].OppPartyInsured, [Personal Injury].InsCo1 FROM tblCase INNER JOIN [Personal Injury] ON tblCase.CaseID = [Personal Injury].CaseID;  | — | yes |
| rptInvoiceComprehensiveAR | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, [Matter and AR].Payment, tblCase.Retainer, tblCase.CaseOpenDate FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID ORDER BY [Matter and AR].Date2;  | — | yes |
| rpt_Compr_InvoiceStmtCur | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprARCur, Report.rptInvoiceComprPymtsARCur, Report.rptInvoiceComprTrustCur | yes |
| rpt_Compr_InvoiceTKExCur | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprARCur, Report.rptInvoiceComprPymtsARCur, Report.rptInvoiceComprTrustCur | yes |
| rptCriminalStatusNotesLog2 | tblNotes | — | yes |
| rpt_MergeInvMatter | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, [Matter and AR].Payment, Nz([Charge],0)-Nz([payment],0) AS Balance, fncRunningDebit([tblcase].[CaseID],[Date2],[MatterID]) AS RunningDebit, fncRunningCredit([tblcase].[CaseID],[Date2],[MatterID]) AS RunningCredit, [RunningDebit]-[RunningCredit] AS RunningBalance, tblCase.Retainer, [RunningBalance]+[retainer] AS RetBal FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID ORDER BY [Matter and AR].MatterID;  | — | yes |
| rpt_Comprehensive_Invoice2 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Trust_Chron_65 | qryTrustEntriesChron65 | — | yes |
| rpt_TimeDetail_Comprehensive2 | qryInvoiceComprehensiveTimeDetail2 | — | yes |
| rpt_TKLessTrust | qryInvoiceAttachRPT1 | Report.rptInvoiceComprehensiveTrust2 | yes |
| rpt_Comprehensive_InvoiceTKEx2 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceADV | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, tblCase.Executor, [TB Time Keeping].ARatClose FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveTrust2, Report.rptInvoiceComprehensiveAR2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_OpenCases | qryCaseListOpen | — | yes |
| rpt_Comprehensive_InvoiceTKEx2S | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceADVS | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, tblCase.Executor, [TB Time Keeping].ARatClose FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveTrust2, Report.rptInvoiceComprehensiveAR2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceTKEx1 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceStmt | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, [TB Time Keeping].AdvBalanceatClose, [TB Time Keeping].TrustatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceTKEx | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Open_Cases | qryTakeOff | — | yes |
| rpt_Comprehensive_InvoiceTKEx1S | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, [TB Time Keeping].Discount, [TB Time Keeping].OutsAdvDue, [TB Time Keeping].AdvFeesBal, [TB Time Keeping].AdvCostBal, [TB Time Keeping].ReplenishBalanceatClose, [TB Time Keeping].ARatClose, tblCase.Executor FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceTKLessTrustCostAR | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID WHERE ((([TB Time Keeping].Bill_ID)=75) AND ((tblCase.CaseID)=14879));  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceTKLessTrustRep | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, tblCase.RetReimbAmount, [TB Time Keeping].Discount FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_Comprehensive_InvoiceTKLessTrustRep2 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClose, tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].TimeNotes, tblCase.Retainer, tblCase.RetReimbAmount, [TB Time Keeping].Discount FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID;  | Report.rptInvoiceComprehensiveAR2, Report.rptInvoiceComprehensiveTrust2, Report.rpt_TimeDetail_Comprehensive2, Report.rptInvoiceComprPymtsAR | yes |
| rpt_TimeDetail_Comprehensive | qryInvoiceComprehensiveTimeDetail | — | yes |
| rpt_File_Folder_Label | qryFileFolderLabel | — | yes |
| rpt_ftrustee_address_label | Bankruptcy | — | yes |
| rpt_MergeInvTimeDetail | tblTimeTableDetail | — | yes |
| Rpt_MergeInvTK | qryInvoiceRPT1 | Report.rpt_MergeInvMatter, Report.rpt_MergeInvTimeDetail | yes |
| rpt_trustee_address_label | Bankruptcy | — | yes |
| rpt_opp_counsel_address_label | tblCase | — | yes |
| rpt_TKExceedsTrust | qryInvoiceAttachRPT1 | — | yes |
| rpt_Trust_Chron_35D | qryTrustEntriesChronRPT35D | — | yes |
| rpt_Trust_Chron_65D | qryTrustEntriesChronRPT65D | — | yes |
| rpt_Trust_Chron_95D | qryTrustEntriesChronRPT95D | — | yes |
| rpt_Trust_Chron_35W | qryTrustEntriesChronRPT35W | — | yes |
| rpt_Trust_Chron_65W | qryTrustEntriesChronRPT65W | — | yes |
| rpt_Trust_Chron_95W | qryTrustEntriesChronRPT95W | — | yes |
| rpt_Trust_Chron_95 | qryTrustEntriesChronRPT95 | — | yes |
| rptBillingTotals | SELECT vwBillingTracker2.Tatty, Sum(vwBillingTracker2.Time_) AS SumOfTime_, Sum(vwBillingTracker2.Billed) AS SumOfBilled, [forms]![frm_Billing_Tracker2]![txtFrom] AS StartDate, [forms]![frm_Billing_Tracker2]![txtTo] AS EndDate FROM vwBillingTracker2 WHERE (((vwBillingTracker2.Tdate) Between [forms]![frm_Billing_Tracker2]![txtFrom] And [forms]![frm_Billing_Tracker2]![txtTo])) GROUP BY vwBillingTracker2.Tatty ORDER BY vwBillingTracker2.Tatty DESC;  | — | yes |
| rptClientNotes | SELECT tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblNotes.NoteDate, tblNotes.NotePerson, tblNotes.NoteDescription, tblNotes.NoteTime, Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNo, tblNotes.IDNotes FROM tblCase INNER JOIN tblNotes ON tblCase.CaseID = tblNotes.CaseID;  | — | yes |
| rptComprehensiveTKStatement | qryInvoiceAttachComp | — | yes |
| rptCriminalStatus | qryCrimStatus | Report.rptCriminalStatusActionNeeded, Report.rptCriminalStatusChargeNos, Report.rptCriminalStatusUpcHrgs, Report.rptCriminalStatusNotesLog2 | yes |
| rptCriminalStatusActionNeeded | SELECT TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, TblActionNeeded.CaseID, TblActionNeeded.DateComp, TblActionNeeded.DateComp1 FROM TblActionNeeded WHERE (((TblActionNeeded.ActionComp)=No));  | — | yes |
| rptCriminalStatusChargeNos | tbl_CtCaseNumbers | — | yes |
| rptCriminalStatusUpcHrgs | SELECT tblHearingDate.CaseID, tblHearingDate.Hearing_Date, tblHearingDate.HearingType, tblHearingDate.HearingTime, tblHearingDate.Verified, tblHearingDate.HrgResult, tblHearingDate.HrgCal, tblHearingDate.ClientPresent, tblHearingDate.Reminder, tblHearingDate.ReminderCheck, tblHearingDate.HearingID FROM tblHearingDate WHERE (((tblHearingDate.Hearing_Date)>Date()));  | — | yes |
| rptInvoiceComprARCur | SELECT * FROM qry_InvoiceAR_curr;  | — | yes |
| rptInvoiceComprehensiveTrust | qryInvoiceComprehensiveTrustCredit | — | yes |
| rptInvoiceComprPymtsAR | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Payment, tblCase.CaseOpenDate, [Matter and AR].OrderNr, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Bill_ID FROM (tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID) INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID WHERE ((([Matter and AR].Date2)<=[Bill Closed Date]) AND (([Matter and AR].Pay_Outlay) Not Like "Adjustment") AND (([Matter and AR].Payment)>0)) ORDER BY [Matter and AR].Date2;  | — | yes |
| rptInvoiceComprPymtsARCur | qry_InvoicePymts_curr | — | yes |
| rptInvoiceComprTrustCur | qryInvoiceComprehensiveTrustCredit | — | yes |
| rptLastWeekIntake | SELECT [TB Intakes].ID, [TB Intakes].[GI Last Name], [TB Intakes].[GI First Name], [TB Intakes].[GI phone], [TB Intakes].[GI Date], [TB Intakes].[GI Practice Area], [TB Intakes].[GI Individual Referrer], [TB Intakes].[GI Comments], [TB Intakes].[GI No Further Action], [TB Intakes].[GI Open], [TB Intakes].[GI Open Date], [TB Intakes].[GI Referral], [TB Intakes].ReasonDintHire, [TB Intakes].FollowUpDate, [TB Intakes].Attorny, [TB Intakes].QuotedFee FROM [TB Intakes] WHERE ((([TB Intakes].[GI Date]) Between getSTDT() And getENDT()));  | — | yes |
| rptPersInjProviderBills | tblPersInjProv | — | yes |
| rptPersInjStatusAction | SELECT TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, TblActionNeeded.CaseID FROM TblActionNeeded WHERE (((TblActionNeeded.ActionComp)=No));  | — | yes |
| rptPersInjStatusDemand | tblPersInjDemand | — | yes |
| rptPersInjStatusLog | tblPersInjLog | — | yes |
| rptPersInjuryStatus | qryPersInjStatus | Report.rptPersInjStatusLog, Report.rptPersInjStatusAction, Report.rptPersInjStatusDemand, Report.rptPersInjProviderBills | yes |
| rptPIStatusSOL | qryAttyTrustAcctsTOff | — | yes |
| rptReceipt | qryReceipt | — | yes |
| rptReceiptC | qryReceipt | — | yes |
| rptReceiptR | tblReceipts | — | yes |
| rptReceiptRec | tblReceipts | — | yes |
| rptReconciliation | tblTakeOffMonth | Report.rpt_Reconciliation sub | yes |
| rptTKReport | qryTimeKeeping | — | yes |
| rptTKReport2 | qryTimeKeeping | — | yes |
| Statement of Trust Account | qryStmtTrustRPT1 | — | yes |

### Queries → referenced tables
| Query | Type | References |
|---|---|---|
| qry_advanced_nonadvanced_payments | select | Matter and AR, tblCase |
| qry_advanced_payments | select | vw_advanced_payments |
| qry_advanced_payments_OLD | select | Matter and AR, tblDropD |
| qry_advanced_totals | select | Matter and AR, tblCase |
| qry_advanced_totals_SUM | select | vw_advanced_totals_SUM |
| qry_advanced_totals_SUM_OLD | select | — |
| qry_caseID_clients | select | TblCase |
| qry_client_names | select | tblCase |
| qry_client_names_TK | select | TblCase |
| qry_CtNames_list_options | select | tblDropD |
| qry_CtType_list_options | select | tblDropD |
| qry_current_invoice | select | vw_current_invoice |
| qry_current_invoice_OLD | select | Billing, Matter and AR |
| qry_disposition_closingSheet | select | Disposition, TblCase |
| qry_file_numbers | select | tblCase |
| qry_find_table_by_field_name | select | tblFields |
| qry_FLChildCustodian_list_options | select | tblDropD |
| qry_FLCompltMethod_list_options | select | tblDropD |
| qry_FLDivorceGrounds_list_options | select | tblDropD |
| qry_FLLengthSeparation_list_options | select | tblDropD |
| qry_FLNOHMethod_list_options | select | tblDropD |
| qry_FLNumberChildren_list_options | select | tblDropD |
| qry_get_MatterID_from_zero_balance | select | — |
| qry_get_time_keeping_numbers | select | TB Time Keeping, tblCase |
| qry_HearingType_list_options | select | tblDropD |
| qry_invoice_comprehensive_trust_acc_cur | select | vw_invoice_comprehensive_trust_acc_cur_unfiltered |
| qry_invoice_comprehensive_trust_acc_cur_OLD | select | Trust Account |
| qry_invoice_comprehensive_trust_acc_cur_unfiltered | select | vw_invoice_comprehensive_trust_acc_cur_unfiltered |
| qry_invoice_comprehensive_trust_acc_cur_unfiltered_old | select | Trust Account |
| qry_InvoiceAR_curr | select | — |
| qry_InvoicePymts_curr | select | — |
| qry_invoices_summary | select | vw_invoices_summary |
| qry_invoices_summary_OLD | select | — |
| qry_invoices_summaryRPT | select | tblDropD |
| qry_last_invoice_sent | select | vw_last_invoice_sent |
| qry_last_invoice_sent_OLD | select | tbl_InvoiceSent |
| qry_LastINV | select | tbl_InvoiceSent |
| qry_matterAR_pay_putlay_list_options | select | tblDropD |
| qry_max_matterID_by_orderNr | select | vw_max_matterID_by_orderNr |
| qry_max_matterID_by_orderNr_OLD | select | Matter and AR |
| qry_orig_atty | select | tblCase |
| qry_orig_atty_filter | select | tblCase |
| qry_OrigAtty_list_options | select | tblDropD |
| qry_RetBalSums_by_PastDue | select | — |
| qry_TA_uncashed_checks | select | TblCase, Trust Account |
| qry_take_off_step2_attorney_sums | select | tblTakeOffMonth |
| qry_take_off_step2_sums | select | vw_take_off_step2_sums |
| qry_take_off_step2_sums_OLD | select | tblTakeOffMonth |
| qry_takeOff_year_month | select | tblTakeOffMonth |
| qry_tblUsers | select | tblAccessType, tblUsers |
| qry_time_table_totals | select | — |
| qry_time_table_totals_atty | select | — |
| qry_time_table_totals_atty_SUM | select | vw_time_table_totals_atty_SUM |
| qry_time_table_totals_atty_SUM_OLD | select | TB Time Keeping |
| qry_time_table_totals_hours | select | — |
| qry_time_table_totals_hours_sum | select | — |
| qry_time_table_totals_SUM | select | vw_time_table_totals_SUM |
| qry_time_table_totals_SUM_OLD | select | TB Time Keeping |
| qry_TimeKeeping_bill_totals | select | TB Time Keeping, tblTimeTableDetail |
| qry_TimeKeeping_CaseID_totals | select | — |
| qry_tmatter_list_options | select | tblDropD |
| qry_trustStatements | select | tblCase |
| qry_uncashed_trust_checks | select | TblCase, Trust Account |
| qryActionNeededAll | select | TblActionNeeded, tblDropD |
| qryActionNeededAll2 | select | TblActionNeeded, tblCase |
| qryActionNeededAll3 | select | TblActionNeeded, tblCase |
| qryActionNeededAllNEW | select | TblActionNeeded, tblCase |
| qryAdvLegalFees | select | Matter and AR |
| qryAdvLegalFeesSum | select | vwAdvLegalFeesSum |
| qryAdvLegalFeesSum_OLD | select | — |
| qryARCredits | select | Matter and AR |
| qryARCreditsSum | select | vwARCreditsSum |
| qryARCreditsSum_OLD | select | — |
| qryAttyTrustAcctsTOff | select | tblCase, tblTakeOff |
| qryBillingTotals | select | tblTimeTableDetail |
| qryBillingTracker | select | tblTimeTableDetail |
| qryBillingTracker2 | select | vwBillingTracker2 |
| qryBillingTracker2_OLD | select | tblCase, tblTimeTableDetail |
| qryBillList | select | TB Time Keeping, TblCase |
| qryCalendarCheck | select | tblCase, tblHearingDate |
| qryCaseIDclientsAll | select | TblCase |
| qryCaseIDclientsClosed | select | TblCase |
| qryCaseIDclientsclosednotscanned | select | TblCase |
| qryCaseList | select | tblCase |
| qryCaseListAll | select | TblCase, tblDropD |
| qryCaseListClosed | select | TblCase, tblDropD |
| qryCaseListOpen | select | vwCaseListOpen |
| qryCaseListOpen_OLD | select | Personal Injury, tblDropD |
| qryCaseSourcesRPT | select | vwCaseSourcesRPT |
| qryCaseSourcesRPT_OLD | select | Disposition, tblCase |
| qryCaseSourcesRPT1 | select | tblDropD |
| qryClosing RPT1 | select | TblCase |
| qryClosingRPT | select | Billing, Disposition, Matter and AR, Trust Account |
| qryCmbCaseClientFile | select | TB Time Keeping, TblCase |
| qryCmbCaseClientFileFamilyLaw | select | Family Law - Divorce, TB Time Keeping |
| qryCostReimb | select | Trust Account |
| qryCostReimbSUM | select | vwCostReimbSUM |
| qryCostReimbSUM_OLD | select | — |
| qryCrimStatus | select | tblCase, tblDropD |
| qryDispoFilter | select | Disposition |
| qryDispos | select | vwDispos |
| qryDispos_OLD | select | Disposition, Personal Injury, tblDropD |
| qryDispos1 | select | Disposition, Personal Injury, tblDropD |
| qryEarnedAdvLegal | select | Trust Account |
| qryEarnedAdvLegalSUM | select | vwEarnedAdvLegalSUM |
| qryEarnedAdvLegalSUM_OLD | select | — |
| qryFamilyLaw | select | Family Law - Divorce, tblCase |
| qryFileFolderLabel | select | tblCase, tblHearingDate |
| qryInvoiceAttachComp | select | tblCase, tblTimeTableDetail |
| qryInvoiceAttachRPT | select | tblCase, tblTimeTableDetail |
| qryInvoiceAttachRPT1 | select | TblCase |
| qryInvoiceComprehensiveTimeDetail | select | TblCase |
| qryInvoiceComprehensiveTimeDetail2 | select | — |
| qryInvoiceComprehensiveTrust | select | vwInvoiceComprehensiveTrust |
| qryInvoiceComprehensiveTrust_OLD | select | tblCase |
| qryInvoiceComprehensiveTrustCredit | select | TB Time Keeping, Trust Account |
| qryInvoiceComprehensiveTrustCredit2 | select | TB Time Keeping |
| qryInvoiceComprehensiveTrustCredit3 | select | TB Time Keeping |
| qryInvoiceComprehensiveTrustCredit4 | select | TB Time Keeping, Trust Account |
| qryInvoiceRPT | select | vwInvoiceRPT |
| qryInvoiceRPT_OLD | select | Billing, Matter and AR |
| qryInvoiceRPT1 | select | vwInvoiceRPT1 |
| qryInvoiceRPT1_OLD | select | tblCase |
| qryInvoiceTrustCostBillDate | select | TB Time Keeping |
| qryMatter | select | vwMatter |
| qryMatter_OLD | select | Matter and AR |
| qryMatterBalanceTotals | select | vwMatterBalanceTotals |
| qryMatterBalanceTotals_OLD | select | — |
| qryMatterSums | select | — |
| qryMergeTest | select | tblCase, tbl_CtCaseNumbers |
| qryNewInvoice_01 | select | — |
| qryNewInvoice_02 | select | — |
| qryNewInvoice02Comp | select | — |
| qryNewTrustComp | select | vwNewTrustComp |
| qryNewTrustComp_OLD | select | tblCase |
| qryOutstandingARRPT | select | Billing |
| qryOutstandingARRPT1 | select | TblCase |
| qryPersInjStatus | select | Personal Injury, tblCase |
| qryReceipt | select | Matter and AR, tblCase |
| qryReconciliation_sumOfBalances | select | — |
| qryReconciliation_sumOfCredit | select | — |
| qryReconciliation_sumOfUnclearedDeposits | select | — |
| qryReconciliationWFBankBalance | select | — |
| qryRunningSum | select | — |
| qrySOL | select | Personal Injury, tblCase |
| qryStmtTrustRPT | select | vwStmtTrustRPT |
| qryStmtTrustRPT_OLD | select | Trust Account, tblCase |
| qryStmtTrustRPT1 | select | vwStmtTrustRPT1 |
| qryStmtTrustRPT1_OLD | select | tblCase |
| qrySumofPayments | select | Matter and AR |
| qryTakeOff | select | — |
| qryTakeOff_A | select | vwTakeOff_A |
| qryTakeOff_A_OLD | select | — |
| qryTakeOff_advanced_AR | select | Matter and AR, tblCase |
| qryTakeOff_cost_hold | select | tblCase |
| qryTakeOff_trust_account | select | vwTakeOff_trust_Account |
| qryTakeOff_trust_account_OLD | select | Trust Account |
| qryTakeOff_unchashed_checks | select | Trust Account, tblCase |
| qryTakeOff_uncleared_deposits | select | Trust Account, tblCase |
| qryTakeOff2 | select | — |
| qryTakeOffAvailBalance | select | — |
| qryTakeOffDate | select | tblTakeOff, tblTakeOffMonth |
| qryTakeOffStep2 | select | vwTakeOffStep2 |
| qryTakeOffStep2_OLD | select | tblCase, tblTakeOff |
| qryTimeKeeping | select | TB Time Keeping, tblCase |
| qryTimeKeepingClosed | select | vwTimeKeepingClosed |
| qryTimeKeepingClosed_Old | select | tblCase |
| qryTimeKeepingOpen | select | vwTimeKeepingOpen |
| qryTimeKeepingOpen_OLD | select | tblCase |
| qryTimeTableRunTot | select | — |
| qryTKClose | select | vwTKClose_A |
| qryTKClose_A | select | vwTKClose_A |
| qryTKClose_A_OLD | select | — |
| qryTKClose_OLD | select | — |
| qryTKClose1 | select | — |
| qryToBeClosed | select | TblCase |
| qryToBeScanned | select | TblCase |
| qryTrustAccount | select | vwTrustAccount |
| qryTrustAccount_OLD | select | Trust Account |
| qryTrustAccountBalanceTotals | select | vwTrustAccountBalanceTotals |
| qryTrustAccountBalanceTotals_OLD | select | — |
| qryTrustCostsExpended | select | Trust Account, tblCase |
| qryTrustCostsExpendedTotals | select | vwTrustCostsExpendedTotals |
| qryTrustCostsExpendedTotals_OLD | select | — |
| qryTrustEntriesChron | select | Trust Account, tblCase |
| qryTrustEntriesChron65 | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT35 | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT35D | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT35W | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT65D | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT65W | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT95 | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT95D | select | Trust Account, tblCase |
| qryTrustEntriesChronRPT95W | select | Trust Account, tblCase |
| qryTrustReportRPT | select | vwTrustReportRPT |
| qryTrustReportRPT_OLD | select | Trust Account, tblCase |
| qryTrustReportRPT1 | select | vwTrustReportRPT1 |
| qryTrustReportRPT1_OLD | select | tblCase |
| qryTrustTotalEarned | select | Trust Account |
| qryTrustTotalEarnedSum | select | vwTrustTotalEarnedSum |
| qryTrustTotalEarnedSum_OLD | select | tblCase |
| qryTTAmount | select | tblTimeTableDetail |
| qryTTAmountAtty | select | tblTimeTableDetail |
| qryTTAmountHours | select | tblTimeTableDetail |
| qryTTAmountHours_SUM | select | tblTimeTableDetail |
| qryTTAmountHours_SUM_byAtty | select | tblTimeTableDetail |
| qryTTAmountHours_SUM_byAtty_TotalCaseID | select | vwTTAmountHours_SUM_byAtty_TotalCaseID |
| qryTTAmountHours_SUM_byAtty_TotalCaseID_OLD | select | TB Time Keeping |
| qryTTAmountHours_TotalCaseID | select | vwTTAmountHours_TotalCaseID |
| qryTTAmountHours_TotalCaseID_OLD | select | TB Time Keeping |
| qryUpcomingHearings | select | tblCase, tblHearingDate |
| qryUpdateattyEmail | update | tblCase |
| Query1 | select | Trust Account, tblCase |
| Sele | select | tblDropD |

### Table → consumers (reverse index)

Which queries, forms, and reports touch each table (forms/reports resolved one hop through their bound query). Tables with no consumers may be lookup/reference data or dead.

| Table | Queries | Forms | Reports |
|---|---|---|---|
| Bankruptcy *(linked)* | — | — | rpt_ftrustee_address_label, rpt_trustee_address_label |
| Billing *(linked)* | qryClosingRPT, qryInvoiceRPT_OLD, qryOutstandingARRPT, qry_current_invoice_OLD | — | — |
| CH13Plans *(linked)* | — | — | — |
| Disposition *(linked)* | qryCaseSourcesRPT_OLD, qryClosingRPT, qryDispoFilter, qryDispos1, qryDispos_OLD, qry_disposition_closingSheet | — | — |
| errMsgs *(linked)* | — | — | — |
| Family Law - Divorce *(linked)* | qryCmbCaseClientFileFamilyLaw, qryFamilyLaw | zfrmFamilyLaw OLD | — |
| Matter and AR *(linked)* | qryARCredits, qryAdvLegalFees, qryClosingRPT, qryInvoiceRPT_OLD, qryMatter_OLD, qryReceipt, qrySumofPayments, qryTakeOff_advanced_AR, qry_advanced_nonadvanced_payments, qry_advanced_payments_OLD, qry_advanced_totals, qry_current_invoice_OLD, qry_max_matterID_by_orderNr_OLD | — | rptReceipt, rptReceiptC |
| Personal Injury *(linked)* | qryCaseListOpen_OLD, qryDispos1, qryDispos_OLD, qryPersInjStatus, qrySOL | frmPersonalInjury2, zfrmPersInjSOL | rptPersInjuryStatus |
| ProofOfClaims *(linked)* | — | — | — |
| TB Intakes *(linked)* | — | Intakes, frmIntakesConflicts | — |
| TB Time Keeping *(linked)* | qryBillList, qryCmbCaseClientFile, qryCmbCaseClientFileFamilyLaw, qryInvoiceComprehensiveTrustCredit, qryInvoiceComprehensiveTrustCredit2, qryInvoiceComprehensiveTrustCredit3, qryInvoiceComprehensiveTrustCredit4, qryInvoiceTrustCostBillDate, qryTTAmountHours_SUM_byAtty_TotalCaseID_OLD, qryTTAmountHours_TotalCaseID_OLD, qryTimeKeeping, qry_TimeKeeping_bill_totals, qry_get_time_keeping_numbers, qry_time_table_totals_SUM_OLD, qry_time_table_totals_atty_SUM_OLD | Time Keeping | rptInvoiceComprTrustCur, rptInvoiceComprehensiveTrust, rptInvoiceComprehensiveTrust2, rptTKReport, rptTKReport2 |
| tbl_CtCaseNumbers *(linked)* | qryMergeTest | — | rptCriminalStatusChargeNos, rpt_CaseNumber_Closing |
| tbl_InvoiceSent *(linked)* | qry_LastINV, qry_last_invoice_sent_OLD | — | — |
| tblAccessType *(linked)* | qry_tblUsers | — | — |
| TblActionNeeded *(linked)* | qryActionNeededAll, qryActionNeededAll2, qryActionNeededAll3, qryActionNeededAllNEW | frmActionNeededAll, frmActionNeededAll2, frmActionNeededAll3 | — |
| tblAttorneys *(linked)* | — | — | — |
| tblCalls *(linked)* | — | frmCalls | — |
| tblCase *(linked)* | Query1, qryActionNeededAll2, qryActionNeededAll3, qryActionNeededAllNEW, qryAttyTrustAcctsTOff, qryBillingTracker2_OLD, qryCalendarCheck, qryCaseList, qryCaseSourcesRPT_OLD, qryCrimStatus, qryFamilyLaw, qryFileFolderLabel, qryInvoiceAttachComp, qryInvoiceAttachRPT, qryInvoiceComprehensiveTrust_OLD, qryInvoiceRPT1_OLD, qryMergeTest, qryNewTrustComp_OLD, qryPersInjStatus, qryReceipt, qrySOL, qryStmtTrustRPT1_OLD, qryStmtTrustRPT_OLD, qryTakeOffStep2_OLD, qryTakeOff_advanced_AR, qryTakeOff_cost_hold, qryTakeOff_unchashed_checks, qryTakeOff_uncleared_deposits, qryTimeKeeping, qryTimeKeepingClosed_Old, qryTimeKeepingOpen_OLD, qryTrustCostsExpended, qryTrustEntriesChron, qryTrustEntriesChron65, qryTrustEntriesChronRPT35, qryTrustEntriesChronRPT35D, qryTrustEntriesChronRPT35W, qryTrustEntriesChronRPT65D, qryTrustEntriesChronRPT65W, qryTrustEntriesChronRPT95, qryTrustEntriesChronRPT95D, qryTrustEntriesChronRPT95W, qryTrustReportRPT1_OLD, qryTrustReportRPT_OLD, qryTrustTotalEarnedSum_OLD, qryUpcomingHearings, qryUpdateattyEmail, qry_advanced_nonadvanced_payments, qry_advanced_totals, qry_client_names, qry_file_numbers, qry_get_time_keeping_numbers, qry_orig_atty, qry_orig_atty_filter, qry_trustStatements | Time Keeping, frmActionNeededAll2, frmCalendarCheck, frmClientsConflict, frmOppPartyConflict, frmTRUSTENTRIESCHRON, frmUpcoming Hearings, frm_trust_summary, zClient Ledger OLD, zfrmFamilyLaw OLD, zfrmPersInjSOL, zfrmPersonalDetailsFamilyLaw | Client_Trust_Accounts_for_Take_Off, rptComprehensiveTKStatement, rptCriminalStatus, rptPIStatusSOL, rptPersInjuryStatus, rptReceipt, rptReceiptC, rptTKReport, rptTKReport2, rpt_File_Folder_Label, rpt_Main_Closing, rpt_Trust_Chron_35, rpt_Trust_Chron_35D, rpt_Trust_Chron_35W, rpt_Trust_Chron_65, rpt_Trust_Chron_65D, rpt_Trust_Chron_65W, rpt_Trust_Chron_95, rpt_Trust_Chron_95D, rpt_Trust_Chron_95W, rpt_address_label, rpt_address_labelEx, rpt_opp_counsel_address_label |
| tblCaseDocuments *(linked)* | — | — | — |
| tblChild *(linked)* | — | — | — |
| tblDocumentRootDirectory *(linked)* | — | — | — |
| tblDocumentTypes *(linked)* | — | — | — |
| tblDropboxLog | — | — | — |
| tblDropboxTokens | — | — | — |
| tblDropD *(linked)* | Sele, qryActionNeededAll, qryCaseListAll, qryCaseListClosed, qryCaseListOpen_OLD, qryCaseSourcesRPT1, qryCrimStatus, qryDispos1, qryDispos_OLD, qry_CtNames_list_options, qry_CtType_list_options, qry_FLChildCustodian_list_options, qry_FLCompltMethod_list_options, qry_FLDivorceGrounds_list_options, qry_FLLengthSeparation_list_options, qry_FLNOHMethod_list_options, qry_FLNumberChildren_list_options, qry_HearingType_list_options, qry_OrigAtty_list_options, qry_advanced_payments_OLD, qry_invoices_summaryRPT, qry_matterAR_pay_putlay_list_options, qry_tmatter_list_options | frmActionNeededAll, frmActionNeededAll3, frmCaseListAll, frmCaseListClosed, frmSourceAnalytics | Accounts Receivable, Case Sources and Revenue, rptCriminalStatus |
| tblFields *(linked)* | qry_find_table_by_field_name | — | — |
| tblFormAccessMapping *(linked)* | — | — | — |
| tblHearingDate *(linked)* | qryCalendarCheck, qryFileFolderLabel, qryUpcomingHearings | frmCalendarCheck, frmUpcoming Hearings | rpt_File_Folder_Label |
| Tblmsgbox *(linked)* | — | — | — |
| tblNotes *(linked)* | — | — | rptCriminalStatusNotesLog2 |
| tblPersInjDemand *(linked)* | — | frmPersInjDemand | rptPersInjStatusDemand |
| tblPersInjLog *(linked)* | — | frmPersInjLog2 | rptPersInjStatusLog |
| tblPersInjProv *(linked)* | — | — | rptPersInjProviderBills |
| tblPrevBank *(linked)* | — | — | — |
| tblReceipts *(linked)* | — | frmReceipt | rptReceiptR, rptReceiptRec |
| tblScans *(linked)* | — | frmScanLocation | — |
| tblTakeOff *(linked)* | qryAttyTrustAcctsTOff, qryTakeOffDate, qryTakeOffStep2_OLD | — | Client_Trust_Accounts_for_Take_Off, rptPIStatusSOL |
| tblTakeOffMonth *(linked)* | qryTakeOffDate, qry_takeOff_year_month, qry_take_off_step2_attorney_sums, qry_take_off_step2_sums_OLD | frmTakeOffTotalFeesCosts | rptReconciliation |
| tblTimeTableDetail *(linked)* | qryBillingTotals, qryBillingTracker, qryBillingTracker2_OLD, qryInvoiceAttachComp, qryInvoiceAttachRPT, qryTTAmount, qryTTAmountAtty, qryTTAmountHours, qryTTAmountHours_SUM, qryTTAmountHours_SUM_byAtty, qry_TimeKeeping_bill_totals | frm_Billing_Tracker | rptComprehensiveTKStatement, rpt_MergeInvTimeDetail |
| tblUsers *(linked)* | qry_tblUsers | frmAddUser | — |
| tblYearMap *(linked)* | — | — | — |
| Trust Account *(linked)* | Query1, qryClosingRPT, qryCostReimb, qryEarnedAdvLegal, qryInvoiceComprehensiveTrustCredit, qryInvoiceComprehensiveTrustCredit4, qryStmtTrustRPT_OLD, qryTakeOff_trust_account_OLD, qryTakeOff_unchashed_checks, qryTakeOff_uncleared_deposits, qryTrustAccount_OLD, qryTrustCostsExpended, qryTrustEntriesChron, qryTrustEntriesChron65, qryTrustEntriesChronRPT35, qryTrustEntriesChronRPT35D, qryTrustEntriesChronRPT35W, qryTrustEntriesChronRPT65D, qryTrustEntriesChronRPT65W, qryTrustEntriesChronRPT95, qryTrustEntriesChronRPT95D, qryTrustEntriesChronRPT95W, qryTrustReportRPT_OLD, qryTrustTotalEarned, qry_TA_uncashed_checks, qry_invoice_comprehensive_trust_acc_cur_OLD, qry_invoice_comprehensive_trust_acc_cur_unfiltered_old, qry_uncashed_trust_checks | frmTRUSTENTRIESCHRON, frm_uncashed_trust_checks | rptInvoiceComprTrustCur, rptInvoiceComprehensiveTrust, rptInvoiceComprehensiveTrust2, rpt_Trust_Chron_35, rpt_Trust_Chron_35D, rpt_Trust_Chron_35W, rpt_Trust_Chron_65, rpt_Trust_Chron_65D, rpt_Trust_Chron_65W, rpt_Trust_Chron_95, rpt_Trust_Chron_95D, rpt_Trust_Chron_95W |
| vw_advanced_payments *(linked)* | qry_advanced_payments | frm_advanced_payments | — |
| vw_advanced_totals_SUM *(linked)* | qry_advanced_totals_SUM | — | — |
| vw_current_invoice *(linked)* | qry_current_invoice | — | New Invoice |
| vw_frm_invoices_summary *(linked)* | — | — | — |
| vw_invoice_comprehensive_trust_acc_cur_unfiltered *(linked)* | qry_invoice_comprehensive_trust_acc_cur, qry_invoice_comprehensive_trust_acc_cur_unfiltered | — | — |
| vw_invoices_summary *(linked)* | qry_invoices_summary | — | — |
| vw_last_invoice_sent *(linked)* | qry_last_invoice_sent | — | — |
| vw_max_matterID_by_orderNr *(linked)* | qry_max_matterID_by_orderNr | — | — |
| vw_rpt_Matter_Closing *(linked)* | — | — | — |
| vw_take_off_step2_sums *(linked)* | qry_take_off_step2_sums | — | — |
| vw_time_table_totals_atty_SUM *(linked)* | qry_time_table_totals_atty_SUM | — | — |
| vw_time_table_totals_SUM *(linked)* | qry_time_table_totals_SUM | — | — |
| vwAdvLegalFeesSum *(linked)* | qryAdvLegalFeesSum | — | — |
| vwARCreditsSum *(linked)* | qryARCreditsSum | — | — |
| vwBillingTracker2 *(linked)* | qryBillingTracker2 | frm_Billing_Tracker2 | — |
| vwCaseListOpen *(linked)* | qryCaseListOpen | frmCaseListOpen | rpt_OpenCases |
| vwCaseSourcesRPT *(linked)* | qryCaseSourcesRPT | — | — |
| vwCostReimbSUM *(linked)* | qryCostReimbSUM | — | — |
| vwDispos *(linked)* | qryDispos | frmDispositions | — |
| vwEarnedAdvLegalSUM *(linked)* | qryEarnedAdvLegalSUM | — | — |
| vwfrmClientLedger *(linked)* | — | — | — |
| vwfrmTakeOffSubForm *(linked)* | — | — | — |
| vwInvoiceComprehensiveTrust *(linked)* | qryInvoiceComprehensiveTrust | — | — |
| vwInvoiceRPT *(linked)* | qryInvoiceRPT | — | — |
| vwInvoiceRPT1 *(linked)* | qryInvoiceRPT1 | — | Invoice, Invoice - No Balance Due, Invoice - Past Due, Invoice2, Rpt_MergeInvTK |
| vwMatter *(linked)* | qryMatter | — | — |
| vwMatterAndAR *(linked)* | — | — | — |
| vwMatterBalanceTotals *(linked)* | qryMatterBalanceTotals | — | — |
| vwNewTrustComp *(linked)* | qryNewTrustComp | — | — |
| vwPILogLatestDate *(linked)* | — | — | — |
| vwStmtTrustRPT *(linked)* | qryStmtTrustRPT | — | — |
| vwStmtTrustRPT1 *(linked)* | qryStmtTrustRPT1 | — | Statement of Trust Account, rpt_Trust_Closing |
| vwTakeOff_A *(linked)* | qryTakeOff_A | — | — |
| vwTakeOff_trust_account *(linked)* | — | — | — |
| vwTakeOffStep2 *(linked)* | qryTakeOffStep2 | frmTakeOffSubForm2, frmTakeOffSubForm3, frmTakeOffSubForm_OLD | — |
| vwTimeKeepingClosed *(linked)* | qryTimeKeepingClosed | frmTimeKeepingClosed | — |
| vwTimeKeepingOpen *(linked)* | qryTimeKeepingOpen | frmTimeKeepingOpen | — |
| vwTimeTableDetail *(linked)* | — | — | — |
| vwTKClose_A *(linked)* | qryTKClose, qryTKClose_A | — | — |
| vwTrustAccount *(linked)* | qryTrustAccount | — | — |
| vwTrustAccountBalanceTotals *(linked)* | qryTrustAccountBalanceTotals | — | — |
| vwTrustAccountTable *(linked)* | — | — | — |
| vwTrustCostsExpendedTotals *(linked)* | qryTrustCostsExpendedTotals | — | — |
| vwTrustReportRPT *(linked)* | qryTrustReportRPT | — | — |
| vwTrustReportRPT1 *(linked)* | qryTrustReportRPT1 | — | — |
| vwTrustTotalEarnedSum *(linked)* | qryTrustTotalEarnedSum | — | — |
| vwTTAmountHours_SUM_byAtty_TotalCaseID *(linked)* | qryTTAmountHours_SUM_byAtty_TotalCaseID | — | — |
| vwTTAmountHours_TotalCaseID *(linked)* | qryTTAmountHours_TotalCaseID | — | — |
| z_PCADataSources | — | — | — |
| z_PCADataSources_TableList | — | — | — |
| z_PCASettings | — | — | — |

## 5. Migration signals

- **Entry point:** `frmHome` (database-properties).
- **Linked tables (88):** Bankruptcy, Billing, CH13Plans, Disposition, errMsgs, Family Law - Divorce, Matter and AR, Personal Injury, ProofOfClaims, TB Intakes, TB Time Keeping, tbl_CtCaseNumbers, tbl_InvoiceSent, tblAccessType, TblActionNeeded, tblAttorneys, tblCalls, tblCase, tblCaseDocuments, tblChild, tblDocumentRootDirectory, tblDocumentTypes, tblDropD, tblFields, tblFormAccessMapping, tblHearingDate, Tblmsgbox, tblNotes, tblPersInjDemand, tblPersInjLog, tblPersInjProv, tblPrevBank, tblReceipts, tblScans, tblTakeOff, tblTakeOffMonth, tblTimeTableDetail, tblUsers, tblYearMap, Trust Account, vw_advanced_payments, vw_advanced_totals_SUM, vw_current_invoice, vw_frm_invoices_summary, vw_invoice_comprehensive_trust_acc_cur_unfiltered, vw_invoices_summary, vw_last_invoice_sent, vw_max_matterID_by_orderNr, vw_rpt_Matter_Closing, vw_take_off_step2_sums, vw_time_table_totals_atty_SUM, vw_time_table_totals_SUM, vwAdvLegalFeesSum, vwARCreditsSum, vwBillingTracker2, vwCaseListOpen, vwCaseSourcesRPT, vwCostReimbSUM, vwDispos, vwEarnedAdvLegalSUM, vwfrmClientLedger, vwfrmTakeOffSubForm, vwInvoiceComprehensiveTrust, vwInvoiceRPT, vwInvoiceRPT1, vwMatter, vwMatterAndAR, vwMatterBalanceTotals, vwNewTrustComp, vwPILogLatestDate, vwStmtTrustRPT, vwStmtTrustRPT1, vwTakeOff_A, vwTakeOff_trust_account, vwTakeOffStep2, vwTimeKeepingClosed, vwTimeKeepingOpen, vwTimeTableDetail, vwTKClose_A, vwTrustAccount, vwTrustAccountBalanceTotals, vwTrustAccountTable, vwTrustCostsExpendedTotals, vwTrustReportRPT, vwTrustReportRPT1, vwTrustTotalEarnedSum, vwTTAmountHours_SUM_byAtty_TotalCaseID, vwTTAmountHours_TotalCaseID — the backend data store must be reproduced/migrated; connect strings are in `schema.json`.
- **Complex fields:** none — schema maps cleanly to standard SQL types.
- **Data macros:** none (expected for `.mdb`; they are `.accdb`-only).
- **VBA red flags:** `createObject`×39, `runSql`×47, `setWarnings`×4, `outputTo`×8, `fileSystem`×20, `eval`×1, `transferSpreadsheet`×1, `transferText`×1 — see `vba/index.json` for the exact procedures.
- ⚠️ CreateObject calls found; external automation needs migration planning.
- ⚠️ File-system access (FileSystemObject/Kill/MkDir) found; reconcile with deployment model.
- ⚠️ 88 linked table(s) detected; back-end connectivity must be reproduced.

## 6. Suggested workflow to build the rewrite plan

1. **Confirm coverage** — read `run_summary.json.completeness`; note any skipped objects as plan risks.
2. **Build the data model** — from `schema.json`; add inferred FKs (§3). Decide the target type for each complex field. This becomes the SQL Server schema + ORM entities.
3. **Catalogue features** — start from `featureCandidates` in the manifest, then refine using the form/report inventory (§4). Each feature → an API surface + UI screen(s).
4. **Wire data sources** — use the data-source map and reverse index (§4) to define API endpoints/entities and see which screens share data.
5. **Extract business logic** — for each VBA procedure / flagged macro / data macro, decide where it lands (validation, service method, DB constraint, background job). Reference exact source under `vba/` and `macros/`.
6. **Resolve high-risk items** (§5) — linked tables, Win32/automation calls, file-system access, dynamic SQL — each needs an explicit migration decision.
7. **Sequence the work** — group into phases (data layer → API → screens → reports), ordered by dependency (entry point `frmHome` and its immediate navigation first).

---
*Generated by the extractor's `migrationGuide` stage from `app_manifest.json` and `run_summary.json`. Regenerate by re-running the extraction.*
