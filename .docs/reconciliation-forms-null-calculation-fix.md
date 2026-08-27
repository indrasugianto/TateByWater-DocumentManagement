# Fix: `txtSMTrust` / `txtSumOfBalances` / `txt_WF_CHK` calculating blank

Companion to [reconciliation-forms-null-calculation-investigation.md](reconciliation-forms-null-calculation-investigation.md) — read that first for root cause and evidence. This doc is the ready-to-apply fix.

**Status: proposed, not yet applied anywhere.** The live `.accdb`/`.accde` isn't in this repo (gitignored, binary) — nothing here can edit it directly. Apply the three blocks below by hand in the Access VBA editor and Query Designer, following the repo's usual rollback convention: the old code is kept as a commented-out `LEGACY` block directly above its replacement, so any block can be rolled back by deleting the new `Private Sub` and uncommenting the old one.

**Scope:** Access-frontend only. `qryTakeOffAvailBalance` (step 1) is a new Access query object built on the existing `qryTakeOff_A` — no SQL Server schema/view changes are needed or made.

**Test before wide deployment.** This is live trust-accounting data for a law firm — if a non-production copy of the `.accdb` is available, apply and verify there first. At minimum, verify in Access before recompiling/redistributing the `.accde`.

---

## Step 1 — New Access query: `qryTakeOffAvailBalance`

`qryTakeOff` can't be reused as a sentinel-CaseID lookup because it silently drops any case whose trust activity nets to $0 (see investigation, H2). This is the same `AvailBalance` formula, without that filter.

In Access: **Create → Query Design → close the "Add Table" dialog → View → SQL View**, paste, then **Save As** `qryTakeOffAvailBalance`:

```sql
SELECT CaseID, [Balance]-Nz([SumOfUnclearedDeposits],0) AS AvailBalance
FROM qryTakeOff_A;
```

Verify it: open it in Datasheet view and confirm `CaseID 12664` now shows `AvailBalance = 0` (a real zero, not a blank row) instead of being absent entirely.

---

## Step 2 — `frmTakeOffReconciliation` → `Form_Load`

Extract reference: `database_assessment/TBCMS/extract/vba/forms/frmTakeOffReconciliation.txt:3300-3323`

Replace the whole procedure with:

```vba
' LEGACY (pre-reconciliation-fix, see .docs/reconciliation-forms-null-calculation-investigation.md)
'Private Sub Form_Load()
'
'    'Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=11946")
'    'Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
'    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
'    Me.txtRLFDeeds = DLookup("AvailBalance", "qryTakeOff", "CaseID=23437")
'    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Nz(Me.txtSMTrust, 0)
'    'Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Me.txtJMTrust - Me.txtSMTrust
'    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
'    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")
'    Me.txt_WF_CHK = Me.txtSumOfBalances + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
'    Me.txtTotalMinusInputBalance = Me.txt_WF_CHK - Me.txt_WF_balance_trust_amount
'    'txtJMTrust = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances")
'    'txtActualJMTrust 'input field
'    Me.txtWFTrustAndJMSMTrust = Me.txtSumOfBalances + Me.txtSMTrust
'    'Me.txtWFTrustAndJMSMTrust = Me.txtSumOfBalances + Me.txtJMTrust + Me.txtSMTrust
'
'    Me.cmdInsertData.Enabled = True
''    If fncReconciliationExists Then
''        cmdInsertData.Enabled = False
''    Else
''        cmdInsertData.Enabled = True
''    End If
'End Sub

Private Sub Form_Load()

    Dim blnReconciliationHasNulls As Boolean

    'Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=11946")
    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=12664")
    Me.txtRLFDeeds = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=23437")
    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Nz(Me.txtSMTrust, 0)
    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")

    blnReconciliationHasNulls = IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits)

    Me.txt_WF_CHK = Nz(Me.txtSumOfBalances, 0) + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
    Me.txtTotalMinusInputBalance = Nz(Me.txt_WF_CHK, 0) - Nz(Me.txt_WF_balance_trust_amount, 0)
    Me.txtWFTrustAndJMSMTrust = Nz(Me.txtSumOfBalances, 0) + Nz(Me.txtSMTrust, 0)

    If blnReconciliationHasNulls Then
        MsgBox "One or more reconciliation values could not be calculated (Null)." & vbCrLf & _
               "Do not rely on the totals on this form until this is resolved.", vbExclamation, "TB CMS"
    End If

    Me.cmdInsertData.Enabled = True
'    If fncReconciliationExists Then
'        cmdInsertData.Enabled = False
'    Else
'        cmdInsertData.Enabled = True
'    End If
End Sub
```

---

## Step 3 — `frmTRUSTENTRIESCHRON` → `Form_Load`

Extract reference: `database_assessment/TBCMS/extract/vba/forms/frmTRUSTENTRIESCHRON.txt:2495-2502`

```vba
' LEGACY (pre-reconciliation-fix, see .docs/reconciliation-forms-null-calculation-investigation.md)
'Private Sub Form_Load()
'    'Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=11946")
'    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
'    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Nz(Me.txtSMTrust, 0)
'    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
'    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")
'    Me.txt_WF_CHK = Me.txtSumOfBalances + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
'End Sub

Private Sub Form_Load()
    'Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=11946")
    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=12664")
    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Nz(Me.txtSMTrust, 0)
    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")

    If IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits) Then
        MsgBox "One or more reconciliation values could not be calculated (Null)." & vbCrLf & _
               "Do not rely on txt_WF_CHK until this is resolved.", vbExclamation, "TB CMS"
    End If

    Me.txt_WF_CHK = Nz(Me.txtSumOfBalances, 0) + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
End Sub
```

---

## Step 4 — `frmTRUSTENTRIESCHRON` → `cmdRequery_Click`

**Don't skip this one.** It's a byte-for-byte third copy of the same broken block, wired to the "Requery" button — miss it and the bug comes right back the moment someone clicks Requery instead of reopening the form.

Extract reference: `database_assessment/TBCMS/extract/vba/forms/frmTRUSTENTRIESCHRON.txt:2481-2489`

```vba
' LEGACY (pre-reconciliation-fix, see .docs/reconciliation-forms-null-calculation-investigation.md)
'Private Sub cmdRequery_Click()
'    If Me.Dirty = True Then Me.Dirty = False
'    Me.Requery
'    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
'    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Nz(Me.txtSMTrust, 0)
'    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
'    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")
'    Me.txt_WF_CHK = Me.txtSumOfBalances + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
'End Sub

Private Sub cmdRequery_Click()
    If Me.Dirty = True Then Me.Dirty = False
    Me.Requery
    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=12664")
    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Nz(Me.txtSMTrust, 0)
    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")

    If IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits) Then
        MsgBox "One or more reconciliation values could not be calculated (Null)." & vbCrLf & _
               "Do not rely on txt_WF_CHK until this is resolved.", vbExclamation, "TB CMS"
    End If

    Me.txt_WF_CHK = Nz(Me.txtSumOfBalances, 0) + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
End Sub
```

---

## Apply & verify

1. **Query Designer:** create `qryTakeOffAvailBalance` (Step 1). Confirm `CaseID=12664` and `CaseID=23437` return a real `0` row.
2. **VBA editor (Alt+F11):** paste Steps 2–4 into the code modules behind `frmTakeOffReconciliation` and `frmTRUSTENTRIESCHRON` respectively, replacing the named procedures exactly (comment the old body, don't delete it).
3. **Debug → Compile** the project; fix any reference errors before proceeding (e.g. if `txtRLFDeeds` or `txtWFTrustAndJMSMTrust` were renamed since this extract was taken).
4. Open both forms and confirm every previously-blank field now shows a real number: `txtSMTrust` and `txtRLFDeeds` should show `$0.00` (not blank), `txtSumOfBalances`/`txt_WF_CHK`/`txtWFTrustAndJMSMTrust` should show large real balances consistent with recent `tblTakeOffMonth` history.
5. Click "Requery" on `frmTRUSTENTRIESCHRON` and re-check `txt_WF_CHK`.
6. If a Null warning box ever pops up during real use, treat it as a real data problem (e.g., `qryReconciliation_sumOfBalances` legitimately returning nothing) — don't dismiss it as another instance of this same bug; investigate before trusting the totals.
7. Save, recompile to `.accde`, redistribute per the normal deployment process for this app.

## Not included in this fix (see investigation doc for why)

- The `CheckCashed=Yes` → `txtSumOfUncashed` naming/semantic inversion (`qryTakeOff_unchashed_checks.sql`) — real, but changes a reported dollar figure; needs separate client sign-off before touching.
- Replacing the hard-coded sentinel CaseIDs (12664/23437/11946) with a config table — recommended longer-term (this is the third time this exact failure mode has hit), but out of scope for this immediate fix.
