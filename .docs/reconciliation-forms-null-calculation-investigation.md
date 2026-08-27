# Investigation: `txtSumofBalances` / `txtSmTrust` / `txt_WF_CHK` calculating blank

**Scope note:** this investigation is entirely inside TBCMS's accounting/trust-reconciliation subsystem. It has no connection to the Dropbox document-management migration that is otherwise this repo's working set (`Dropbox-Migration/*`, `.docs/dropbox-migration-plan.md`). No migration file was touched.

## Summary

The root cause is a **hard-coded sentinel `CaseID` (12664) whose trust balance has drained to exactly $0.00**, which causes it to be silently excluded by a business-logic `WHERE` filter in `qryTakeOff` that the code reuses (inappropriately) as a lookup source for that sentinel value. `DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")` therefore returns `Null`, and `Me.txtSMTrust` renders blank — confirmed live against production today. Whether that `Null` then cascades into `txtSumOfBalances` and `txt_WF_CHK` depends on a second, independently confirmed defect: the code Nz-guards that propagation in one place but not another, and the guard is inconsistent between what today's extract shows and what the client pasted — strong evidence the client's deployed `.accde` build is older/different than the source this investigation read. A related, definitely-still-live defect (present in **every** copy of this code, today's extract included) is that `txt_WF_CHK`'s own formula never Nz-guards `txtSumOfBalances`, so the same failure mode will resurface even after the immediate CaseID problem is patched, the moment any upstream aggregate is legitimately `Null`.

This is not a one-off: the identical pattern already broke a sister sentinel lookup (CaseID 11946, "JM/Dale") years ago, and whoever hit it then simply commented the line out rather than fixing the underlying design flaw (see Evidence, H2). The investigation also turned up a fourth blank field the client didn't report — `txtWFTrustAndJMSMTrust` ("Combined Trust" on `frmTakeOffReconciliation`) — which follows from the same `Null` `txtSMTrust` via its own, separate unguarded-arithmetic line, confirmed live in `tblTakeOffMonth` history (flat `$0.00` for the last 10 months).

## Evidence walkthrough

### H2 — Hard-coded CaseID: CONFIRMED root cause (live-verified, version-independent)

- The client's `Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")` is real code, present verbatim in both forms' actual `Form_Load`:
  - `database_assessment/TBCMS/extract/vba/forms/frmTakeOffReconciliation.txt:3304`
  - `database_assessment/TBCMS/extract/vba/forms/frmTRUSTENTRIESCHRON.txt:2497` (and again in `cmdRequery_Click` at line 2484)
- `qryTakeOff` carries a business filter that excludes any case whose trust activity nets to zero this period (`database_assessment/TBCMS/extract/queries/qryTakeOff.sql:3`):
  ```sql
  WHERE ((([Balance]-Nz([SumOfUnclearedDeposits],0))<>0))
     OR ((([Balance]+Nz([SumOfUncashedChecks],0)-Nz([SumOfUnclearedDeposits],0))<>0))
     OR (((qryTakeOff_A.SumOfUnclearedDeposits)<>0))
     OR (((qryTakeOff_A.SumOfUncashedChecks)<>0))
     OR (((qryTakeOff_A.SumOfTotal)<>0))
  ```
  `qryTakeOff_A` (`database_assessment/TBCMS/extract/queries/qryTakeOff_A.sql`) is an unfiltered passthrough of the SQL Server view `dbo.vwTakeOff_A`.
- **Live query against production (`tbf-cms`/`TateBywater`, read-only)**, CaseID 12664 in `dbo.vwTakeOff_A`:
  ```
  CaseID  Balance  SumOfUncashedChecks  SumOfUnclearedDeposits  SumOfTotal
  12664   .0000    NULL                 NULL                    NULL
  ```
  Applying `qryTakeOff`'s exact `WHERE` clause (translated to T-SQL — Jet/ACE and T-SQL share the same three-valued `NULL` comparison semantics, so the translation is faithful) against CaseID 12664 returns **zero rows**: every term is either `FALSE` (the zero-balance arithmetic terms) or `NULL` (`NULL <> 0` on the three genuinely-`NULL` columns), and `FALSE OR NULL` is `NULL`, which a `WHERE` clause treats as excluding the row. This directly and empirically confirms `DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")` returns `Null` today.
  - Underlying trust ledger for 12664 confirms the balance is genuinely, exactly zero, not a data-quality artifact: `dbo.vwTakeOff_trust_account` shows `SumOfDebit = SumOfCredit = 300.08`, `Balance = .0000`.
- **This has happened before, for the sister sentinel account, and was worked around rather than fixed.** The line immediately above in both forms is commented out:
  ```vba
  'Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=11946")
  ```
  Live query confirms **why**: CaseID 11946 is also at `Balance = .0000` / all-`NULL` activity columns in `vwTakeOff_A` today, and `tblCase` shows it was **closed 2011-12-31**. CaseID 12664 (label on the form: "TD (Somerf)") maps to `tblCase.Last_Name = 'Sommerfeld (Occidental)'`, **closed 2013-01-31**. Both are long-closed cases being reused as fixed trust-sub-ledger markers — a fragile pattern that fails the instant the marked sub-account's balance nets to exactly zero, which is exactly what happened to 12664.
  - A third instance of the same pattern is in `frmTakeOffReconciliation` only (not `frmTRUSTENTRIESCHRON`): `Me.txtRLFDeeds = DLookup("AvailBalance", "qryTakeOff", "CaseID=23437")` (line 3305). Live-checked: CaseID 23437 (`tblCase.Last_Name = 'Deeds Jan-Aug 2019'`) is **also** `Balance = .0000` / all-`NULL` today, so `txtRLFDeeds` is silently blank right now too — not reported by the client, but the same defect, live.
- **Historical timeline, from `tblTakeOffMonth`** (the table `cmdInsertData_Click` persists these fields into — `database_assessment/TBCMS/extract/vba/forms/frmTakeOffReconciliation.txt:3494-3500`, mapping `Me.txtSMTrust`→`SomBalance`, `Me.txtSumOfBalances`→`[WF Balance]`, `Me.txt_WF_CHK`→`WFplusuncashed`):
  ```
  TakeOffDate   WF Balance    SomBalance   WFplusuncashed
  2026-08-03    948,156.35    .0000        988,884.30
  2026-07-01    1,081,708.00  .0000        1,112,048.78
  ...
  2025-09-02    803,772.21    .0000        836,738.28
  2025-08-02    856,654.88    100.0000     886,005.97
  2025-07-01    1,142,030.25  100.0000     1,155,268.96
  2025-06-02    1,176,496.00  100.0000     1,195,673.21
  ```
  `SomBalance` (`txtSMTrust`) has been persisted as exactly `0.00` in **every monthly reconciliation row since 2025-09-02** — a full year — after sitting at a residual `$100.00` through mid-2025. This dates the onset precisely and confirms it is not new. **Important secondary finding:** `[WF Balance]` and `WFplusuncashed` have continued to show large, plausible, non-zero figures every month during that same window, including the most recent row (2026-08-03). Since `cmdInsertData_Click`'s `INSERT` wraps every field in `Nz(...,0)` (line 3497, 3500), a blank `txtSumOfBalances`/`txt_WF_CHK` on the form at insert time would have persisted as `0.00`, not a large number — so whoever has been running the monthly "Insert Data" workflow has, every month for a year, seen a blank/zero `txtSMTrust` but a correctly-computed `txtSumOfBalances`/`txt_WF_CHK`. See H1 below for why that matters.

**A fourth blank field, not in the client's report, follows directly from this same root cause with no version-drift assumption needed:** `frmTakeOffReconciliation.txt:3314`,
```vba
Me.txtWFTrustAndJMSMTrust = Me.txtSumOfBalances + Me.txtSMTrust
```
is unguarded on **both** operands. With `txtSMTrust = Null` (confirmed above), this control is blank today regardless of which build is running — `912134.55 + Null = Null` under today's extract, `Null + Null = Null` under the client's pasted variant. Its on-form label is "Combined Trust" (`Label23`, `frmTakeOffReconciliation.txt:705`), a visible, prominently-placed header field, not a hidden control — the client may simply not have listed it.

**`tblTakeOffMonth` history confirms this independently.** `cmdInsertData_Click` persists it as `Nz(txtWFTrustAndJMSMTrust, 0)` → column `CombinedTrust` (`frmTakeOffReconciliation.txt:3498`):
```
TakeOffDate   CombinedTrust   SomBalance   [WF Balance]
2026-08-03    .0000           .0000        948,156.35
2026-07-01    .0000           .0000        1,081,708.00
...
2025-11-03    .0000           .0000        977,536.41
2025-10-01    .0000           .0000        840,149.13
2025-09-02    803,772.21      .0000        803,772.21
2025-08-02    856,754.88      100.0000     856,654.88
2025-07-01    1,142,130.25    100.0000     1,142,030.25
2025-06-02    1,176,596.00    100.0000     1,176,496.00
```
`CombinedTrust` has been persisted as flat `0.00` in every row from **2025-10-01 onward** (10 straight months), while `[WF Balance]` continued showing large real numbers in the same rows — exactly the signature of `Me.txtSumOfBalances + Me.txtSMTrust` evaluating to `Null` because `txtSMTrust` is `Null`, then getting zero-masked by the `INSERT`'s `Nz()` wrap. The **2025-09-02 row is a sharp corroborating data point**: `CombinedTrust` there exactly equals `[WF Balance]` (both `803,772.21`) — i.e. `txtSMTrust` was still contributing a real `$0.00` that month, not `Null`. That pins the onset of CaseID 12664 falling out of `qryTakeOff` to **between 2025-09-02 and 2025-10-01**, tighter than the `SomBalance` column alone could show (which was already `0.0000` by 2025-09-02, since it can't distinguish "computed to real zero" from "Null zero-masked by `Nz()` on insert" the way the `CombinedTrust` comparison can).

**Verdict: CONFIRMED.** This is the root cause of `txtSMTrust` (and, via `txtWFTrustAndJMSMTrust`, "Combined Trust") going blank, live-verified against production data, and dated to between 2025-09-02 and 2025-10-01.

### H1 — Missing `Nz()` on the final subtraction/addition: CONFIRMED mechanism, but split into two findings

**What today's extract actually contains** (not what the client pasted — see H4 for the diff):
```vba
' frmTakeOffReconciliation.txt:3304-3310, frmTRUSTENTRIESCHRON.txt:2497-2501 (Form_Load)
' and frmTRUSTENTRIESCHRON.txt:2484-2488 (cmdRequery_Click) — same body, three sites total
Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Nz(Me.txtSMTrust, 0)
Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")
Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")
Me.txt_WF_CHK = Me.txtSumOfBalances + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0)
```

1. **The `txtSumOfBalances` line already Nz-guards `Me.txtSMTrust`** in every extracted copy of this code (both forms, all three call sites). This is *not* what the client pasted (client's version has bare `- Me.txtSMTrust`, no `Nz`).
2. **The `txt_WF_CHK` line does *not* Nz-guard `Me.txtSumOfBalances`**, in any of the three sites, in today's extract. This is confirmed present, unconditionally, right now. Per Access/VBA Null-propagation rules, `Null + x` and `Null - x` both evaluate to `Null` regardless of whether `x` itself is wrapped in `Nz()` — the `Nz()` on the *other* operands does nothing to protect against a `Null` `txtSumOfBalances`. This is standard, well-documented VBA/Jet behavior, not project-specific.

**Live-data check of what these two findings predict for *today's* code, given today's data:**
- `DLookup("SumOfBalance", "qryReconciliation_sumOfBalances")` → **912,134.55** (live query: `SELECT COUNT(*), SUM(Balance) FROM dbo.vwTakeOff_trust_account` → 6,679 rows, `SumBalance = 912134.55`). Non-null.
- `DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")` → **25,363.83** (live: `SELECT COUNT(*), SUM(Credit) FROM dbo.[Trust Account] WHERE CheckCashed=1` → 36 rows, non-null).
- `DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")` → **7,006.64** (live: same pattern on `DepCleared=1` → 9 rows, non-null).
- With `Nz(Me.txtSMTrust, 0)` in place (today's extract), `txtSumOfBalances = 912134.55 - 0 = 912134.55` — a real number, **not blank**. `txt_WF_CHK = 912134.55 + Nz(25363.83,0) - Nz(7006.64,0) = 930,491.74` — also **not blank**.

**This is inconsistent with the client's report of all three fields being blank, and it is corroborated by the `tblTakeOffMonth` history above showing real `[WF Balance]`/`WFplusuncashed` numbers every month while `SomBalance` sat at `0.00`.** Two independent signals — the client's own pasted code (missing the `Nz` on the `txtSumOfBalances` line) and the client's reported symptom (all three fields blank, not just `txtSMTrust`) — both point the same direction and are both inconsistent with today's extracted source. The most parsimonious explanation is that **the `.accde` the client is actually running is a different build than the source this extract was taken from today** (`app_manifest.json` shows this extract was pulled `2026-08-27` from `TB CMS.SQL.accdb`, a working copy, not necessarily the exact build compiled into the client's production `.accde`). Per-form/per-user Access frontends compiled at different times is exactly the deployment shape this app has (see `CLAUDE.md`: "Frontend = per-user Access `.accde`"). A copy/paste transcription slip on the client's part is possible but less likely to independently produce a symptom set that matches so precisely — it would need to be a "matching" slip that happens to reproduce this exact code branch's behavior.

**Verdict:**
- The `txt_WF_CHK` → `txtSumOfBalances` Nz-gap is **CONFIRMED present in source today**, at all three call sites, and is a live landmine regardless of which build the client runs: the day `DLookup("SumOfBalance", "qryReconciliation_sumOfBalances")` legitimately returns `Null` (e.g., the trust-account view ever returns zero rows), `txt_WF_CHK` will go blank even with all other guards in place.
- The `txtSumOfBalances` → `txtSMTrust` Nz-gap is **the client's likely actual live experience** (their pasted code lacks it), but is **already patched in the extract read for this investigation** — meaning the fix may already be sitting unshipped in a `.accdb` working copy, or the client's build simply predates it. This needs a build/version check (see Open Questions) rather than a code change, since the code-as-extracted already has this specific guard.

### H3 — Control-name mismatch across forms: RULED OUT

All six controls referenced by the pasted code (`txtSumOfBalances`, `txtSMTrust`, `txtSumOfUncashed`, `txtSumOfUnclearDeposits`, `txt_WF_CHK`, `txtJMTrust`) exist as `TextBox` controls on **both** forms, confirmed via two independent sources:
- Structured control extract: `database_assessment/TBCMS/extract/forms/frmTakeOffReconciliation.json` and `frmTRUSTENTRIESCHRON.json` (both list all six by name/type).
- Raw form-definition text: `frmTakeOffReconciliation.txt` (`txtSumOfBalances` L725, `txtSMTrust` L1118, `txtSumOfUncashed` L933, `txtSumOfUnclearDeposits` L973, `txt_WF_CHK` L1012, `txtJMTrust` L765) and `frmTRUSTENTRIESCHRON.txt` (`txt_WF_CHK` L1230, `txtSumOfBalances` L2145, `txtSumOfUncashed` L2165, `txtSumOfUnclearDeposits` L2186, `txtJMTrust` L2259, `txtSMTrust` L2282).

No `Me.` reference in `Form_Load` targets a non-existent control on either form, so there is no unhandled-runtime-error path here. Confirmed further: neither form's `Form_Load` contains an `On Error` handler (the only `On Error` in either module is in `CaseNum_Click`, unrelated — `frmTakeOffReconciliation.txt:3393`, `frmTRUSTENTRIESCHRON.txt:2363`), so if a runtime error *did* occur it would surface as a visible Access error dialog, not a silent blank field. Since all `DLookup` calls in `Form_Load` resolve to valid, existing query objects with valid criteria syntax (confirmed by successfully reproducing every one of them against production above), there is no runtime-error condition present — the `Null` values are legitimate "no matching row" / "no matching data" DLookup returns, not exceptions.

### H4 — Do the two forms share one `Form_Load`, and does the client's paste match "frmTrustEntriesChron" twice: RESOLVED

They are **two independent `Form_Load` procedures**, one per form module, not a shared/copy-pasted-into-both-by-reference body:
- `frmTakeOffReconciliation.txt:3300-3323` — 24 lines, sets `txtSMTrust`, `txtRLFDeeds`, `txtSumOfBalances`, `txtSumOfUncashed`, `txtSumOfUnclearDeposits`, `txt_WF_CHK`, `txtTotalMinusInputBalance`, `txtWFTrustAndJMSMTrust`, and `cmdInsertData.Enabled = True`.
- `frmTRUSTENTRIESCHRON.txt:2495-2502` — 8 lines, sets only `txtSMTrust`, `txtSumOfBalances`, `txtSumOfUncashed`, `txtSumOfUnclearDeposits`, `txt_WF_CHK`.

The client's pasted snippet is a **near-exact match for `frmTRUSTENTRIESCHRON`'s `Form_Load`** — same six statements, same comments (including the literal typo `'Not cashed! please check qyery !`), same hard-coded CaseIDs, same commented-out `txtJMTrust` line — differing only in the one missing `Nz()` discussed under H1. It is **not** a match for `frmTakeOffReconciliation`'s `Form_Load`, which has four additional lines the client never pasted. So: the "both labeled frmTrustEntriesChron" oddity in the bug report is explained by the client pasting `frmTRUSTENTRIESCHRON`'s code twice (once, presumably, intending to show `frmTakeOffReconciliation`'s) rather than by any code-sharing between the forms. **`frmTakeOffReconciliation`'s actual `Form_Load` code was never actually shown by the client** — worth flagging back to them, since its extra logic (`txtWFTrustAndJMSMTrust`, `txtTotalMinusInputBalance`) is untested by anything in this investigation beyond the live DLookup checks already covered.

Also worth noting: `frmTRUSTENTRIESCHRON`'s `cmdRequery_Click` (L2481-2489) contains a **third, byte-for-byte copy** of this same block, wired to the "Requery" button. Any fix must be applied there too, or the bug returns the moment a user clicks Requery instead of reopening the form.

### H5 — Other Null sources in the query chain: RULED OUT for today's data, but architecture is worth flagging

- `qryReconciliation_sumOfBalances` (`SELECT Sum(Balance) FROM qryTakeOff_trust_account`, no `WHERE`, no `GROUP BY`) can only return `Null` if the entire 6,679-row `vwTakeOff_trust_account` view were empty or every `Balance` were `Null`. Live-checked: not the case today (912,134.55).
- `qryReconciliation_sumOfCredit` / `qryReconciliation_sumOfUnclearedDeposits` (`Sum(...)` over `qryTakeOff_unchashed_checks` / `qryTakeOff_uncleared_deposits`, both `HAVING`-filtered on `[Trust Account].CheckCashed`/`DepCleared`) would return `Null` only if literally zero rows in `[Trust Account]` matched `CheckCashed=Yes` / `DepCleared=Yes` firm-wide. Live-checked: 36 and 9 matching rows respectively today, both non-null sums.
- No `WHERE` clause anywhere in this chain references an open form/control value or a date range — these are all firm-wide aggregates, evaluated fully at `Form_Load` time with no dependency on other controls being populated first. So there's no "control not yet loaded" ordering hazard.
- **Separate, real defect independently flagged by the client's own code comment** — not a Null source, filed separately per scope: `qryTakeOff_unchashed_checks.sql` computes `SumOfCredit` `HAVING CheckCashed = Yes` (i.e., **cashed** checks), but the result feeds a field named `txtSumOfUncashed`. The client's inline comment (`'Not cashed! please check qyery !`) already flags this semantic inversion. It does not cause a blank value (confirmed non-null above) — it's a correctness/labeling defect, not part of this bug's root cause.
- One more Null-source, low-impact and out of the client's reported list: `txtBBSum`'s `ControlSource` (`frmTakeOffReconciliation.txt:1663-1664`, control is hidden — `Visible = NotDefault`) is `=Sum(Nz([Bankbalance],0))-[txtJMTrust]-[txtSMTrust]`. Since `txtJMTrust` is never assigned (its `Form_Load` line is permanently commented out), this hidden control's expression is permanently `Null` regardless of the CaseID issue. Noted for completeness; it isn't visible to users and wasn't reported.

**Verdict:** No other Null source is currently active in this chain. The two confirmed mechanisms (H2 root cause + H1 propagation gap) fully account for the reported symptoms.

## Ranking

1. **H2 — hard-coded `CaseID=12664` dropped out of the `WHERE`-filtered `qryTakeOff`** — CONFIRMED, live-verified, present in every build/version of this code (the `DLookup` call itself is identical in the client's paste and today's extract). This is the actual root cause and the one concrete, dateable, reproducible defect.
2. **H1 — missing `Nz()` around `txtSumOfBalances` in the `txt_WF_CHK` line** — CONFIRMED present in source today at all three call sites; this is the propagation mechanism that turns a single blank field (`txtSMTrust`) into a cascading blank (`txt_WF_CHK`), and will keep doing so for any future legitimate-Null upstream aggregate even after H2 is fixed.
3. **H1b — missing `Nz()` around `txtSMTrust` in the `txtSumOfBalances` line** — this one specific gap is *not* present in today's extract (it's already guarded), which is inconsistent with the client seeing `txtSumOfBalances` blank too. Best explanation: build/version drift between the client's deployed `.accde` and the source read here (see Open Questions) — not a code defect to fix in *this* source, since it's already fixed here.
4. **H3 (control mismatch) and H5 (other Null sources)** — RULED OUT, both confirmed via primary source + live data.
5. **H4** — resolved as a labeling artifact in the client's bug report, not a code-sharing bug.

## Recommended fix

**Do not simply wrap everything in `Nz(...,0)`.** This is a trust-account reconciliation tool for a law firm; silently defaulting a missing trust figure to `$0.00` makes a reconciliation *look* balanced when the real number was never computed — masking exactly the kind of discrepancy this tool exists to catch. Fix the source of the `Null`, not just its propagation.

**1. Stop using the `WHERE`-filtered `qryTakeOff` as the lookup source for fixed sentinel CaseIDs.** `qryTakeOff_A` (`database_assessment/TBCMS/extract/queries/qryTakeOff_A.sql`) is the unfiltered passthrough already sitting underneath it — but it does not carry the `AvailBalance` column (that's computed only inside `qryTakeOff`'s own `SELECT` list, `qryTakeOff.sql:1`; `vwTakeOff_A`/`qryTakeOff_A` only expose `Balance` and `SumOfUnclearedDeposits` separately — confirmed via `database_assessment/TBCMS/extract/tables/vwTakeOff_A.md` column list, no `AvailBalance` field). Two concrete options, in order of preference:
   - **Preferred:** add a small unfiltered helper query, e.g. `qryTakeOffAvailBalance`:
     ```sql
     SELECT CaseID, [Balance]-Nz([SumOfUnclearedDeposits],0) AS AvailBalance
     FROM qryTakeOff_A;
     ```
     Then repoint the three sentinel `DLookup`s at it instead of `qryTakeOff`:
     ```vba
     Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=12664")
     Me.txtRLFDeeds = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=23437")
     ' and, if ever un-commented: Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=11946")
     ```
     This leaves `qryTakeOff`'s business filter (used by the main reconciliation grid and `rptReconciliation`'s trigger chain) completely untouched.
   - **Alternative (no new query object):** compute it inline from two `DLookup`s against `qryTakeOff_A` directly:
     ```vba
     Me.txtSMTrust = Nz(DLookup("Balance", "qryTakeOff_A", "CaseID=12664"), 0) _
                    - Nz(DLookup("SumOfUnclearedDeposits", "qryTakeOff_A", "CaseID=12664"), 0)
     ```
     Slightly more duplication of the `AvailBalance` formula (DRY concern), but avoids a new query object.
   - Apply the same change to `txtRLFDeeds`'s `CaseID=23437` lookup (`frmTakeOffReconciliation.txt:3305`) — it is failing today for the identical reason and simply hasn't been reported yet.

   Verify with the same live-query technique used in this investigation: after the fix, `DLookup("AvailBalance", "qryTakeOffAvailBalance", "CaseID=12664")` should return `0` (i.e. `0 - Nz(Null,0)`, since `Balance=0` and `SumOfUnclearedDeposits` is `Null` for this case) rather than `Null` — restoring `txtSMTrust` to a real `$0.00` (a legitimate reconciled-out balance) instead of a blank.

   **This one change is sufficient to resolve all three fields the client reported, on either build.** Once `txtSMTrust` returns a real `$0.00` instead of `Null`: under the client's own (unguarded) pasted code, `txtSumOfBalances = 912134.55 - 0.00` computes normally, and `txt_WF_CHK` follows from it — no dependency on first resolving which `.accde` build is deployed (Open Question 1). The build question still matters for confirming *why* the client saw three blanks instead of one, and for deciding whether fix #2 needs to ship to their machine too, but it does not block or gate fix #1's effectiveness.

**2. Close the `txt_WF_CHK` Nz-gap, at all three call sites**, so a *future* legitimate Null upstream doesn't silently blank the field again — but pair it with a visible warning rather than silent zeroing, given the trust-accounting context:
   ```vba
   If IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits) Then
       MsgBox "One or more reconciliation values could not be calculated (Null). " & _
              "Do not rely on txt_WF_CHK until this is resolved.", vbExclamation, "TB CMS"
   End If
   Me.txt_WF_CHK = Nz(Me.txtSumOfBalances, 0) + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0)
   ```
   Apply to all three sites: `frmTakeOffReconciliation.txt:3310` (`Form_Load`), `frmTRUSTENTRIESCHRON.txt:2501` (`Form_Load`), `frmTRUSTENTRIESCHRON.txt:2488` (`cmdRequery_Click`). Missing the third site means the bug reappears the instant someone clicks the "Requery" button instead of reopening the form.

   `frmTakeOffReconciliation`-only: apply the identical guard-plus-warning treatment to the two downstream unguarded lines that chain off `txtSumOfBalances`/`txt_WF_CHK` on that form —
   ```vba
   Me.txtWFTrustAndJMSMTrust = Nz(Me.txtSumOfBalances, 0) + Nz(Me.txtSMTrust, 0)   ' L3314, "Combined Trust" — confirmed blank today
   Me.txtTotalMinusInputBalance = Nz(Me.txt_WF_CHK, 0) - Nz(Me.txt_WF_balance_trust_amount, 0)   ' L3311
   ```
   Skipping these leaves "Combined Trust" blank on the form header even after `txt_WF_CHK` itself is fixed, since it has its own independent Nz-gap (see H2's "fourth blank field" finding).

**3. Do not touch `qryTakeOff_unchashed_checks`'s `CheckCashed=Yes` filter without separate client sign-off.** It's a real semantic-inversion defect (feeds a field literally named `txtSumOfUncashed`), already self-flagged in the code's own comment, but it is not part of this bug (confirmed non-null, not blank) and changing it changes a reported dollar figure — get explicit confirmation of intended business meaning before touching it.

**4. Consider a structural fix, not just a patch**, given this is the *third* time a hard-coded sentinel CaseID has broken this way (11946 already dead, 12664 breaking now, 23437 already broken and unreported): replace the fixed-CaseID pattern with a small config table (e.g., a `tblTrustSentinelAccounts(Label, CaseID)` row set) that's looked up instead of literal IDs baked into VBA, so a future zero-balance case doesn't require another code change to work around. This is a design recommendation, not required to close the immediate bug.

## Open questions

1. **Which `.accde` build is the client actually running, and does it already contain `Nz(Me.txtSMTrust, 0)` on the `txtSumOfBalances` line?** This determines whether the client is currently seeing all three fields blank (matches their unguarded paste) or would, after today's-extract behavior, see only `txtSMTrust` blank. The fastest way to check: have the client note whether `txtSumOfBalances` currently shows a large number (~$900K–$1M) or is truly blank. If it's a real number, only fix #1 above is needed for them; if it's blank too, their build needs to be reconciled/recompiled from the same source this extract came from (or fix #2 needs to ship to their build specifically).
2. **Is the person who has been running the monthly "Insert Data" workflow (producing the real `[WF Balance]`/`WFplusuncashed` figures in `tblTakeOffMonth` every month through 2026-08-03) a different user/machine than whoever reported this bug?** If so, this is very likely simple per-user `.accde` version drift, consistent with the architecture described in `CLAUDE.md` ("Frontend = per-user Access `.accde`").
3. **Should the `tblTakeOffMonth` rows already persisted with `SomBalance = 0.00` since 2025-09-02** (a year of monthly reconciliation records where the Sommerfeld sentinel was silently zeroed instead of showing its true, already-reconciled $0 balance — which in this specific case happens to be numerically correct today, but was blank-masked-to-zero rather than genuinely computed) **be reviewed/corrected**, or is `$0.00` in fact accurate for that account for that whole period? Live data suggests it likely is accurate (the account has been at exactly `$0.00` balance the whole time), but this should be confirmed by whoever owns the reconciliation, not assumed by this investigation.
4. `frmTakeOffReconciliation`'s fuller `Form_Load` (the version the client never actually pasted) also sets `txtTotalMinusInputBalance = Me.txt_WF_CHK - Me.txt_WF_balance_trust_amount` (`frmTakeOffReconciliation.txt:3311`), unguarded. It inherits `txt_WF_CHK`'s blank/real-number status directly and should get the same fix as item 2 below. (`txtWFTrustAndJMSMTrust` itself is no longer an open question — see H2's "fourth blank field" finding above; it's confirmed, not speculative.)
