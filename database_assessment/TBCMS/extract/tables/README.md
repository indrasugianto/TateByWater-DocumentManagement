# Table Definitions

One definition file per table (91 tables). Structured source of truth: `../schema.json`.

| Table | Columns | Primary key | Linked | Rows |
|-------|---------|-------------|--------|------|
| [Bankruptcy](Bankruptcy.md) | 35 | BankruptcyID | yes | — |
| [Billing](Billing.md) | 10 | ID | yes | — |
| [CH13Plans](CH13Plans.md) | 9 | IDCH13Plans | yes | — |
| [Disposition](Disposition.md) | 14 | DispoID | yes | — |
| [errMsgs](errMsgs.md) | 4 | ID | yes | — |
| [Family Law - Divorce](Family_Law_-_Divorce.md) | 61 | ID | yes | — |
| [Matter and AR](Matter_and_AR.md) | 12 | MatterID | yes | — |
| [Personal Injury](Personal_Injury.md) | 57 | ID | yes | — |
| [ProofOfClaims](ProofOfClaims.md) | 9 | IDProofOfClaims | yes | — |
| [TB Intakes](TB_Intakes.md) | 21 | ID | yes | — |
| [TB Time Keeping](TB_Time_Keeping.md) | 27 | Bill_ID | yes | — |
| [tbl_CtCaseNumbers](tbl_CtCaseNumbers.md) | 8 | CtCaseNoID | yes | — |
| [tbl_InvoiceSent](tbl_InvoiceSent.md) | 12 | InvoiceSentID | yes | — |
| [tblAccessType](tblAccessType.md) | 5 | ID | yes | — |
| [TblActionNeeded](TblActionNeeded.md) | 9 | ActionNeededID | yes | — |
| [tblAttorneys](tblAttorneys.md) | 9 | AttysID | yes | — |
| [tblCalls](tblCalls.md) | 21 | CallID | yes | — |
| [tblCase](tblCase.md) | 104 | CaseID | yes | — |
| [tblCaseDocuments](tblCaseDocuments.md) | 5 | CaseDocumentID | yes | — |
| [tblChild](tblChild.md) | 4 | Child_ID | yes | — |
| [tblDocumentRootDirectory](tblDocumentRootDirectory.md) | 10 | DocumentRootDirectoryID | yes | — |
| [tblDocumentTypes](tblDocumentTypes.md) | 6 | DocumentTypeID | yes | — |
| [tblDropD](tblDropD.md) | 8 | DropID | yes | — |
| [tblFields](tblFields.md) | 6 | — | yes | — |
| [tblFormAccessMapping](tblFormAccessMapping.md) | 4 | ID | yes | — |
| [tblHearingDate](tblHearingDate.md) | 12 | HearingID | yes | — |
| [Tblmsgbox](Tblmsgbox.md) | 3 | — | yes | — |
| [tblNotes](tblNotes.md) | 7 | IDNotes | yes | — |
| [tblPersInjDemand](tblPersInjDemand.md) | 5 | PIDemandID | yes | — |
| [tblPersInjLog](tblPersInjLog.md) | 6 | PersInjLogID | yes | — |
| [tblPersInjProv](tblPersInjProv.md) | 9 | PIProviderID | yes | — |
| [tblPrevBank](tblPrevBank.md) | 6 | IDPrevBank | yes | — |
| [tblReceipts](tblReceipts.md) | 13 | ReceiptID | yes | — |
| [tblScans](tblScans.md) | 5 | ScansID | yes | — |
| [tblTakeOff](tblTakeOff.md) | 31 | TakeOffID | yes | — |
| [tblTakeOffMonth](tblTakeOffMonth.md) | 41 | TakeOffMonthID | yes | — |
| [tblTimeTableDetail](tblTimeTableDetail.md) | 8 | Time_ID | yes | — |
| [tblUsers](tblUsers.md) | 4 | ID | yes | — |
| [tblYearMap](tblYearMap.md) | 4 | — | yes | — |
| [Trust Account](Trust_Account.md) | 14 | TrustAccountID | yes | — |
| [vw_advanced_payments](vw_advanced_payments.md) | 17 | CaseID | yes | — |
| [vw_advanced_totals_SUM](vw_advanced_totals_SUM.md) | 3 | CaseID | yes | — |
| [vw_current_invoice](vw_current_invoice.md) | 22 | CaseID | yes | — |
| [vw_frm_invoices_summary](vw_frm_invoices_summary.md) | 19 | CaseID | yes | — |
| [vw_invoice_comprehensive_trust_acc_cur_unfiltered](vw_invoice_comprehensive_trust_acc_cur_unfiltered.md) | 6 | CaseID | yes | — |
| [vw_invoices_summary](vw_invoices_summary.md) | 18 | CaseID | yes | — |
| [vw_last_invoice_sent](vw_last_invoice_sent.md) | 2 | CaseID | yes | — |
| [vw_max_matterID_by_orderNr](vw_max_matterID_by_orderNr.md) | 3 | CaseID | yes | — |
| [vw_rpt_Matter_Closing](vw_rpt_Matter_Closing.md) | 13 | MatterID | yes | — |
| [vw_take_off_step2_sums](vw_take_off_step2_sums.md) | 13 | TakeOffMonthID | yes | — |
| [vw_time_table_totals_atty_SUM](vw_time_table_totals_atty_SUM.md) | 3 | CaseID | yes | — |
| [vw_time_table_totals_SUM](vw_time_table_totals_SUM.md) | 2 | CaseID | yes | — |
| [vwAdvLegalFeesSum](vwAdvLegalFeesSum.md) | 3 | CaseID | yes | — |
| [vwARCreditsSum](vwARCreditsSum.md) | 2 | CaseID | yes | — |
| [vwBillingTracker2](vwBillingTracker2.md) | 16 | Time_ID | yes | — |
| [vwCaseListOpen](vwCaseListOpen.md) | 18 | CaseID | yes | — |
| [vwCaseSourcesRPT](vwCaseSourcesRPT.md) | 11 | CaseID | yes | — |
| [vwCostReimbSUM](vwCostReimbSUM.md) | 2 | CaseID | yes | — |
| [vwDispos](vwDispos.md) | 20 | CaseID | yes | — |
| [vwEarnedAdvLegalSUM](vwEarnedAdvLegalSUM.md) | 2 | CaseID | yes | — |
| [vwfrmClientLedger](vwfrmClientLedger.md) | 105 | CaseID | yes | — |
| [vwfrmTakeOffSubForm](vwfrmTakeOffSubForm.md) | 144 | TakeOffID | yes | — |
| [vwInvoiceComprehensiveTrust](vwInvoiceComprehensiveTrust.md) | 6 | CaseID | yes | — |
| [vwInvoiceRPT](vwInvoiceRPT.md) | 13 | CaseID | yes | — |
| [vwInvoiceRPT1](vwInvoiceRPT1.md) | 117 | CaseID | yes | — |
| [vwMatter](vwMatter.md) | 7 | MatterID | yes | — |
| [vwMatterAndAR](vwMatterAndAR.md) | 15 | MatterID | yes | — |
| [vwMatterBalanceTotals](vwMatterBalanceTotals.md) | 2 | CaseID | yes | — |
| [vwNewTrustComp](vwNewTrustComp.md) | 117 | CaseID | yes | — |
| [vwPILogLatestDate](vwPILogLatestDate.md) | 2 | ID | yes | — |
| [vwStmtTrustRPT](vwStmtTrustRPT.md) | 19 | CaseID | yes | — |
| [vwStmtTrustRPT1](vwStmtTrustRPT1.md) | 117 | CaseID | yes | — |
| [vwTakeOff_A](vwTakeOff_A.md) | 20 | CaseID | yes | — |
| [vwTakeOff_trust_account](vwTakeOff_trust_account.md) | 4 | CaseID | yes | — |
| [vwTakeOffStep2](vwTakeOffStep2.md) | 132 | TakeOffID | yes | — |
| [vwTimeKeepingClosed](vwTimeKeepingClosed.md) | 117 | CaseID | yes | — |
| [vwTimeKeepingOpen](vwTimeKeepingOpen.md) | 108 | CaseID | yes | — |
| [vwTimeTableDetail](vwTimeTableDetail.md) | 8 | Time_ID | yes | — |
| [vwTKClose_A](vwTKClose_A.md) | 25 | CaseID | yes | — |
| [vwTrustAccount](vwTrustAccount.md) | 6 | TrustAccountID | yes | — |
| [vwTrustAccountBalanceTotals](vwTrustAccountBalanceTotals.md) | 2 | CaseID | yes | — |
| [vwTrustAccountTable](vwTrustAccountTable.md) | 16 | TrustAccountID | yes | — |
| [vwTrustCostsExpendedTotals](vwTrustCostsExpendedTotals.md) | 2 | CaseID | yes | — |
| [vwTrustReportRPT](vwTrustReportRPT.md) | 12 | CaseID | yes | — |
| [vwTrustReportRPT1](vwTrustReportRPT1.md) | 12 | CaseID | yes | — |
| [vwTrustTotalEarnedSum](vwTrustTotalEarnedSum.md) | 2 | CaseID | yes | — |
| [vwTTAmountHours_SUM_byAtty_TotalCaseID](vwTTAmountHours_SUM_byAtty_TotalCaseID.md) | 3 | CaseID | yes | — |
| [vwTTAmountHours_TotalCaseID](vwTTAmountHours_TotalCaseID.md) | 2 | CaseID | yes | — |
| [z_PCADataSources](z_PCADataSources.md) | 9 | PCADataSourceName, ApplicationStatus | — | — |
| [z_PCADataSources_TableList](z_PCADataSources_TableList.md) | 5 | PCADataSourceName, ConnectAs | — | — |
| [z_PCASettings](z_PCASettings.md) | 3 | INISection, INIKey | — | — |
