# Table Definitions

One definition file per table (93 tables). Structured source of truth: `../schema.json`.

| Table | Columns | Primary key | Linked | Rows |
|-------|---------|-------------|--------|------|
| [Bankruptcy](Bankruptcy.md) | 35 | BankruptcyID | yes | 211 |
| [Billing](Billing.md) | 10 | ID | yes | 2043 |
| [CH13Plans](CH13Plans.md) | 9 | IDCH13Plans | yes | 191 |
| [Disposition](Disposition.md) | 14 | DispoID | yes | 7497 |
| [errMsgs](errMsgs.md) | 4 | ID | yes | 2 |
| [Family Law - Divorce](Family_Law_-_Divorce.md) | 61 | ID | yes | 28 |
| [Matter and AR](Matter_and_AR.md) | 12 | MatterID | yes | 22953 |
| [Personal Injury](Personal_Injury.md) | 57 | ID | yes | 444 |
| [ProofOfClaims](ProofOfClaims.md) | 9 | IDProofOfClaims | yes | 903 |
| [TB Intakes](TB_Intakes.md) | 21 | ID | yes | 1724 |
| [TB Time Keeping](TB_Time_Keeping.md) | 27 | Bill_ID | yes | 8743 |
| [tbl_CtCaseNumbers](tbl_CtCaseNumbers.md) | 8 | CtCaseNoID | yes | 9689 |
| [tbl_InvoiceSent](tbl_InvoiceSent.md) | 12 | InvoiceSentID | yes | 12529 |
| [tblAccessType](tblAccessType.md) | 5 | ID | yes | 6 |
| [TblActionNeeded](TblActionNeeded.md) | 9 | ActionNeededID | yes | 5347 |
| [tblAttorneys](tblAttorneys.md) | 9 | AttysID | yes | 0 |
| [tblCalls](tblCalls.md) | 21 | CallID | yes | 13171 |
| [tblCase](tblCase.md) | 104 | CaseID | yes | 12099 |
| [tblCaseDocuments](tblCaseDocuments.md) | 5 | CaseDocumentID | yes | 27226 |
| [tblChild](tblChild.md) | 4 | Child_ID | yes | 28 |
| [tblDocumentRootDirectory](tblDocumentRootDirectory.md) | 10 | DocumentRootDirectoryID | yes | 1 |
| [tblDocumentTypes](tblDocumentTypes.md) | 6 | DocumentTypeID | yes | 29 |
| [tblDropboxLog](tblDropboxLog.md) | 5 | LogID | — | 59 |
| [tblDropboxTokens](tblDropboxTokens.md) | 8 | TokenID | — | 1 |
| [tblDropD](tblDropD.md) | 8 | DropID | yes | 245 |
| [tblFields](tblFields.md) | 6 | — | yes | 0 |
| [tblFormAccessMapping](tblFormAccessMapping.md) | 4 | ID | yes | 5 |
| [tblHearingDate](tblHearingDate.md) | 12 | HearingID | yes | 14670 |
| [Tblmsgbox](Tblmsgbox.md) | 3 | — | yes | 4 |
| [tblNotes](tblNotes.md) | 7 | IDNotes | yes | 8006 |
| [tblPersInjDemand](tblPersInjDemand.md) | 5 | PIDemandID | yes | 1309 |
| [tblPersInjLog](tblPersInjLog.md) | 6 | PersInjLogID | yes | 6036 |
| [tblPersInjProv](tblPersInjProv.md) | 9 | PIProviderID | yes | 1940 |
| [tblPrevBank](tblPrevBank.md) | 6 | IDPrevBank | yes | 66 |
| [tblReceipts](tblReceipts.md) | 13 | ReceiptID | yes | 1451 |
| [tblScans](tblScans.md) | 5 | ScansID | yes | 4678 |
| [tblTakeOff](tblTakeOff.md) | 31 | TakeOffID | yes | 45719 |
| [tblTakeOffMonth](tblTakeOffMonth.md) | 41 | TakeOffMonthID | yes | 147 |
| [tblTimeTableDetail](tblTimeTableDetail.md) | 8 | Time_ID | yes | 61287 |
| [tblUsers](tblUsers.md) | 4 | ID | yes | 50 |
| [tblYearMap](tblYearMap.md) | 4 | — | yes | 35 |
| [Trust Account](Trust_Account.md) | 14 | TrustAccountID | yes | 44799 |
| [vw_advanced_payments](vw_advanced_payments.md) | 17 | CaseID | yes | 542 |
| [vw_advanced_totals_SUM](vw_advanced_totals_SUM.md) | 3 | CaseID | yes | 1180 |
| [vw_current_invoice](vw_current_invoice.md) | 22 | CaseID | yes | 22953 |
| [vw_frm_invoices_summary](vw_frm_invoices_summary.md) | 19 | CaseID | yes | 5641 |
| [vw_invoice_comprehensive_trust_acc_cur_unfiltered](vw_invoice_comprehensive_trust_acc_cur_unfiltered.md) | 6 | CaseID | yes | 44799 |
| [vw_invoices_summary](vw_invoices_summary.md) | 18 | CaseID | yes | 5641 |
| [vw_last_invoice_sent](vw_last_invoice_sent.md) | 2 | CaseID | yes | 1900 |
| [vw_max_matterID_by_orderNr](vw_max_matterID_by_orderNr.md) | 3 | CaseID | yes | 5641 |
| [vw_rpt_Matter_Closing](vw_rpt_Matter_Closing.md) | 13 | MatterID | yes | 22953 |
| [vw_take_off_step2_sums](vw_take_off_step2_sums.md) | 13 | TakeOffMonthID | yes | 111 |
| [vw_time_table_totals_atty_SUM](vw_time_table_totals_atty_SUM.md) | 3 | CaseID | yes | 5738 |
| [vw_time_table_totals_SUM](vw_time_table_totals_SUM.md) | 2 | CaseID | yes | 2839 |
| [vwAdvLegalFeesSum](vwAdvLegalFeesSum.md) | 3 | CaseID | yes | 925 |
| [vwARCreditsSum](vwARCreditsSum.md) | 2 | CaseID | yes | 165 |
| [vwBillingTracker2](vwBillingTracker2.md) | 16 | Time_ID | yes | 61287 |
| [vwCaseListOpen](vwCaseListOpen.md) | 18 | CaseID | yes | 1990 |
| [vwCaseSourcesRPT](vwCaseSourcesRPT.md) | 11 | CaseID | yes | 7000 |
| [vwCostReimbSUM](vwCostReimbSUM.md) | 2 | CaseID | yes | 1072 |
| [vwDispos](vwDispos.md) | 20 | CaseID | yes | 6932 |
| [vwEarnedAdvLegalSUM](vwEarnedAdvLegalSUM.md) | 2 | CaseID | yes | 591 |
| [vwfrmClientLedger](vwfrmClientLedger.md) | 105 | CaseID | yes | 12099 |
| [vwfrmTakeOffSubForm](vwfrmTakeOffSubForm.md) | 144 | TakeOffID | yes | 52463 |
| [vwInvoiceComprehensiveTrust](vwInvoiceComprehensiveTrust.md) | 6 | CaseID | yes | 50212 |
| [vwInvoiceRPT](vwInvoiceRPT.md) | 13 | CaseID | yes | 22953 |
| [vwInvoiceRPT1](vwInvoiceRPT1.md) | 117 | CaseID | yes | 22953 |
| [vwMatter](vwMatter.md) | 7 | MatterID | yes | 22953 |
| [vwMatterAndAR](vwMatterAndAR.md) | 15 | MatterID | yes | 22953 |
| [vwMatterBalanceTotals](vwMatterBalanceTotals.md) | 2 | CaseID | yes | 5641 |
| [vwNewTrustComp](vwNewTrustComp.md) | 117 | CaseID | yes | 50212 |
| [vwPILogLatestDate](vwPILogLatestDate.md) | 2 | ID | yes | 373 |
| [vwStmtTrustRPT](vwStmtTrustRPT.md) | 19 | CaseID | yes | 50212 |
| [vwStmtTrustRPT1](vwStmtTrustRPT1.md) | 117 | CaseID | yes | 50212 |
| [vwTakeOff_A](vwTakeOff_A.md) | 20 | CaseID | yes | 12101 |
| [vwTakeOff_trust_account](vwTakeOff_trust_account.md) | 4 | CaseID | yes | 6686 |
| [vwTakeOffStep2](vwTakeOffStep2.md) | 132 | TakeOffID | yes | 52463 |
| [vwTimeKeepingClosed](vwTimeKeepingClosed.md) | 117 | CaseID | yes | 95 |
| [vwTimeKeepingOpen](vwTimeKeepingOpen.md) | 108 | CaseID | yes | 1121 |
| [vwTimeTableDetail](vwTimeTableDetail.md) | 8 | Time_ID | yes | 61287 |
| [vwTKClose_A](vwTKClose_A.md) | 25 | CaseID | yes | 5643 |
| [vwTrustAccount](vwTrustAccount.md) | 6 | TrustAccountID | yes | 44799 |
| [vwTrustAccountBalanceTotals](vwTrustAccountBalanceTotals.md) | 2 | CaseID | yes | 6686 |
| [vwTrustAccountTable](vwTrustAccountTable.md) | 16 | TrustAccountID | yes | 44799 |
| [vwTrustCostsExpendedTotals](vwTrustCostsExpendedTotals.md) | 2 | CaseID | yes | 1957 |
| [vwTrustReportRPT](vwTrustReportRPT.md) | 12 | CaseID | yes | 50212 |
| [vwTrustReportRPT1](vwTrustReportRPT1.md) | 12 | CaseID | yes | 12135 |
| [vwTrustTotalEarnedSum](vwTrustTotalEarnedSum.md) | 2 | CaseID | yes | 6475 |
| [vwTTAmountHours_SUM_byAtty_TotalCaseID](vwTTAmountHours_SUM_byAtty_TotalCaseID.md) | 3 | CaseID | yes | 5738 |
| [vwTTAmountHours_TotalCaseID](vwTTAmountHours_TotalCaseID.md) | 2 | CaseID | yes | 2839 |
| [z_PCADataSources](z_PCADataSources.md) | 9 | PCADataSourceName, ApplicationStatus | — | 2 |
| [z_PCADataSources_TableList](z_PCADataSources_TableList.md) | 5 | PCADataSourceName, ConnectAs | — | 88 |
| [z_PCASettings](z_PCASettings.md) | 3 | INISection, INIKey | — | 4 |
