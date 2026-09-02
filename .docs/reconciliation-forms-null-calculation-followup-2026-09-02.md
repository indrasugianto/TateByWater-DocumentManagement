# Follow-up: "Trust Account (Chron)" / Requery Null-reconciliation popup (2026-09-02)

Companion to [reconciliation-forms-null-calculation-investigation.md](reconciliation-forms-null-calculation-investigation.md) (root-cause investigation) and [reconciliation-forms-null-calculation-fix.md](reconciliation-forms-null-calculation-fix.md) (proposed fix, both committed 2026-08-27, `83f72af`). Read those first — this doc does not repeat their evidence, only what's new since then.

## What the user reported today

Clicking the **"Trust Account (Chron)"** tile on the Admin page, or the **"Requery"** button on the **"Trust Account Entries"** page, now pops:

> **TB CMS**
> One or more reconciliation values could not be calculated (Null).
> Do not rely on txt_WF_CHK until this is resolved.

Both actions land on the same form. `cmdTrustAccountChron_Click` on the Admin form does `DoCmd.openform "frmTrustEntriesChron"` (`database_assessment/TBCMS/extract/vba/forms/frmHomeAdmin.txt:6274-6275`), and the "Trust Account Entries" page's "Requery" button is `frmTRUSTENTRIESCHRON`'s own `cmdRequery_Click`. So both repro paths hit `frmTRUSTENTRIESCHRON`, matching the object name convention already established in the prior investigation (H4).

> **Update, 2026-09-02 15:15 UTC — extract refreshed, everything below confirmed against deployed source.** The user re-extracted the Access objects after this investigation's first pass. The refreshed extract confirms the full original fix (Steps 1–4) is deployed on **both** forms, and settles the two questions this doc had left open. Findings 1 and 3 are annotated inline where the refresh changed the answer; Finding 2's conclusion is unchanged and is now confirmed directly rather than inferred.

## Finding 1 — this popup is the guard code the prior fix doc proposed, already deployed live

The wording is a byte-for-byte match for the `MsgBox` the fix doc specified for **`frmTRUSTENTRIESCHRON` specifically** (not `frmTakeOffReconciliation`, which uses different wording — "Do not rely on **the totals on this form**" — see [reconciliation-forms-null-calculation-fix.md:79-81](reconciliation-forms-null-calculation-fix.md#L79-L81) vs [:117-118](reconciliation-forms-null-calculation-fix.md#L117-L118) / [:154-155](reconciliation-forms-null-calculation-fix.md#L154-L155)):

```vba
If IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits) Then
    MsgBox "One or more reconciliation values could not be calculated (Null)." & vbCrLf & _
           "Do not rely on txt_WF_CHK until this is resolved.", vbExclamation, "TB CMS"
End If
Me.txt_WF_CHK = Nz(Me.txtSumOfBalances, 0) + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0)
```

That block is the fix doc's **Step 3** (`Form_Load`) and **Step 4** (`cmdRequery_Click`) for `frmTRUSTENTRIESCHRON` — exactly the two entry points the user just triggered. Someone has already implemented this part of the proposed fix live.

*(Originally inferred from the popup text, because the then-current `2026-08-27` extract predated the deployment and showed only the unguarded pre-fix bodies.)* **Confirmed directly by the `2026-09-02` re-extract:**

- `database_assessment/TBCMS/extract/vba/forms/frmTRUSTENTRIESCHRON.txt:2485-2499` (`cmdRequery_Click`) and `:2516-2529` (`Form_Load`) — both carry the guard verbatim, each with its pre-fix body preserved as a commented `LEGACY` block above it (`:2474-2483`, `:2506-2514`), per repo convention.
- `database_assessment/TBCMS/extract/vba/forms/frmTakeOffReconciliation.txt:3319-3347` — `frmTakeOffReconciliation` **also** has the guard deployed, via `blnReconciliationHasNulls` (`:3330`), with the "Do not rely on the totals on this form" wording (`:3337-3338`).

So the fix doc's Steps 1–4 were all applied, on both forms.

## Finding 2 — the popup's exact dollar figure pins the firing branch to `txtSumOfUnclearDeposits`, and confirms the app is pointed at `awsql2022dev`

Per the user's steer, I queried `awsql2022dev` / `TateByWater` directly (the environment CLAUDE.md names for Test) for the three `DLookup` sources the deployed guard checks:

```sql
SELECT COUNT(*) AS Cnt, SUM(Balance) AS SumBalance FROM dbo.vwTakeOff_trust_account;
-- Cnt 6686, SumBalance 701356.29   (feeds qryReconciliation_sumOfBalances)

SELECT COUNT(*) AS Cnt, SUM(Credit) AS SumCredit FROM dbo.[Trust Account] WHERE CheckCashed=1;
-- Cnt 48, SumCredit 60247.91      (feeds qryReconciliation_sumOfCredit)

SELECT COUNT(*) AS Cnt, SUM(Debit) AS SumDebit FROM dbo.[Trust Account] WHERE DepCleared=1;
-- Cnt 0, SumDebit NULL             (feeds qryReconciliation_sumOfUnclearedDeposits)
```

Working through the deployed formula with these values and the confirmed-still-live sentinel-`CaseID` defect (Finding 3 below, `txtSMTrust` = `Null`):

```
txtSumOfBalances = 701356.29 - Nz(Null, 0)            = 701356.29        (non-null)
txtSumOfUncashed = 60247.91                            = 60247.91        (non-null)
txtSumOfUnclearDeposits = DLookup(... zero rows ...)   = Null            (NULL — zero rows matched DepCleared=Yes)

txt_WF_CHK = Nz(701356.29,0) + Nz(60247.91,0) - Nz(Null,0) = 761,604.20
```

**`761,604.20` is the exact figure shown in the red "WELLS FARGO" box on the "Trust Account Entries" screenshot.** That's a cent-exact match on an 8-digit computed value against live data pulled independently from `awsql2022dev` today — strong enough to draw two conclusions with high confidence:

1. **The screenshots are from a build pointed at `awsql2022dev`/`TateByWater`**, not production — this database's numbers are what's on screen. (Also consistent with production having materially different figures: the original investigation saw `6,679` rows / `912,134.55` on `dbo.vwTakeOff_trust_account` against production; today's dev-mirror check sees `6,686` rows / `701,356.29` — a different, unrelated dataset, as expected for a separate environment.)
2. **The guard's third branch — `IsNull(Me.txtSumOfUnclearDeposits)` — is the one firing.** `txtSumOfBalances` and `txtSumOfUncashed` both compute to real, non-null numbers on this data; only `txtSumOfUnclearDeposits` is `Null`, and it's `Null` because **zero rows in `[Trust Account]` currently have `DepCleared = Yes`** (`database_assessment/TBCMS/extract/queries/qryTakeOff_uncleared_deposits.sql` is `GROUP BY ... HAVING DepCleared=Yes`; zero matching rows → `qryReconciliation_sumOfUnclearedDeposits`'s `Sum()` over an empty set → `Null`).

**This is the root cause of today's popup.** It is not something the original investigation (which was scoped to `txtSMTrust`/`CaseID 12664`) checked or ruled out, since `DepCleared` wasn't implicated by any symptom at the time.

**This is not a transient data-timing gap — `DepCleared` has never once been `True`, for any row, in the entire table's history:**

```sql
SELECT DepCleared, COUNT(*) AS Cnt FROM dbo.[Trust Account] GROUP BY DepCleared;
-- DepCleared=0: 44,799 rows. No DepCleared=1 rows exist at all.
SELECT MIN(TDate), MAX(TDate), COUNT(*) FROM dbo.[Trust Account];
-- 1899-12-30 through 2026-09-02, 44,799 rows total.
```

So `qryTakeOff_uncleared_deposits`'s `HAVING DepCleared=Yes` is *guaranteed* to return zero rows and `Null` every time it runs, on this environment, indefinitely — not "currently zero, will fill in later." The guard will keep firing every load/Requery until something changes structurally, not on its own.

**Likely explanation — the same label/field-semantics inversion already flagged (and left unfixed) for the checks side:** the on-screen column header directly above the `DepCleared` checkbox reads **"Uncl Deposit"** (`Label112`, `database_assessment/TBCMS/extract/vba/forms/frmTRUSTENTRIESCHRON.txt:707`, positioned at `Left=14415` — directly over the `DepCleared` checkbox at `Left=14895`, same column band). So the UI presents this box to staff as "check this if the deposit is *uncleared*," but the query reads a checked box (`DepCleared=Yes`) as "cleared." This is the deposit-side mirror of the exact pattern the original investigation flagged for `CheckCashed`/"Uncashed Check" (H5: `qryTakeOff_unchashed_checks.sql` filters `CheckCashed=Yes`, feeding a field named `txtSumOfUncashed`, already self-flagged by the code's own comment `'Not cashed! please check qyery !'` — and explicitly left unfixed pending client sign-off, per the fix doc's "Not included in this fix" list).

The two fields aren't symmetric in practice, though: `CheckCashed` **is** actively used (48 of 44,799 rows are `1`), just possibly filtered backwards. `DepCleared` has **never** been used at all — either staff have never used the "Uncl Deposit" checkbox in this app's entire history, or deposit-side reconciliation happens through a different mechanism entirely. One candidate: `Reconciled` (a third checkbox/column on the same table) is heavily used — `44,599` of `44,799` rows are `Reconciled=1`, `200` are `0` — but nothing in the extracted queries filters on it for reconciliation purposes today (`qryTrustEntriesChron.sql` and others just pass it through as a display column). Whether `Reconciled` is meant to be the real "has this cleared" signal for deposits, or is unrelated (e.g. set by the monthly `cmdInsertData` close routine for a different purpose), is a business-logic question this investigation can't answer from the code alone.

## Finding 3 — the original sentinel-`CaseID` defect (H2) is **fixed**, and was never what caused this popup

**Resolved by the `2026-09-02` re-extract.** The sentinel fix (fix doc Step 1) is fully deployed:

- `database_assessment/TBCMS/extract/queries/qryTakeOffAvailBalance.sql` now exists, exactly as prescribed:
  ```sql
  SELECT CaseID, [Balance]-Nz([SumOfUnclearedDeposits],0) AS AvailBalance FROM qryTakeOff_A;
  ```
- All sentinel `DLookup`s are repointed to it — `frmTRUSTENTRIESCHRON.txt:2488` (`cmdRequery_Click`) and `:2518` (`Form_Load`); `frmTakeOffReconciliation.txt:3324` (`CaseID=12664`) and `:3325` (`CaseID=23437`), plus the commented-out `11946` line at `:3323`.
- Live check through the new query's logic confirms it returns real zeros, not `Null`:
  ```sql
  SELECT CaseID, Balance - ISNULL(SumOfUnclearedDeposits,0) AS AvailBalance
  FROM dbo.vwTakeOff_A WHERE CaseID IN (12664,23437);
  -- 12664 -> .0000    23437 -> .0000   (rows present, values real)
  ```

So `txtSMTrust` and `txtRLFDeeds` now resolve to a genuine `$0.00` rather than blanking, and the "Combined Trust" cascade the original investigation documented is closed. **This also retires this doc's earlier open question** about whether the `DLookup` had been repointed — it had; the arithmetic simply couldn't discriminate, since `701,356.29 - Nz(Null,0)` and `701,356.29 - 0` are identical.

The underlying *data* condition is unchanged and always will be — the sentinel cases still net to zero:

```sql
SELECT CaseID, Balance, SumOfUncashedChecks, SumOfUnclearedDeposits, SumOfTotal
FROM dbo.vwTakeOff_A WHERE CaseID IN (12664, 23437, 11946);
```
```
CaseID Balance SumOfUncashedChecks SumOfUnclearedDeposits SumOfTotal
12664  .0000   NULL                 NULL                    NULL
11946  .0000   NULL                 NULL                    NULL
23437  .0000   NULL                 NULL                    NULL
```

All three still net to `$0.00` with `NULL` activity columns, so `qryTakeOff`'s business `WHERE` filter still excludes them — which is exactly why the fix routed the sentinel lookups around that filter via `qryTakeOffAvailBalance` instead of trying to change the filter. The fix is correct and holding; nothing further is needed here.

Even before the repoint, this was never the popup's cause: `txtSumOfBalances` does `... - Nz(Me.txtSMTrust, 0)`, so a `Null` `txtSMTrust` was absorbed before the guard ever evaluated it.

So H2 is a **real, still-open, but currently latent** defect — worth fixing (per the original fix doc), but fixing it will not silence today's popup. Do not lead with it when responding to this specific report.

## How to stop the popup — the options, and what each one costs

Because `DepCleared` is structurally always `False`, there is no "wait for the data to come good" path. Something has to change. In order of how much they touch reported dollar figures:

**Option 1 — Decide what "Uncl Deposit" is supposed to mean, then fix the query to match (recommended first step, but needs sign-off).** If the intent is genuinely "deposits recorded but not yet cleared by the bank" — the standard bank-reconciliation adjustment — then `HAVING DepCleared = Yes` is filtering the wrong population and should be `= No` (or the column should be `Reconciled`, or something else entirely). Any of those would return a real, non-`Null` number and permanently stop the popup. **But it changes `txt_WF_CHK` — a reported trust-account figure — so it needs explicit accounting/business sign-off first**, exactly as the original investigation required for the identical `CheckCashed` case. Note the two are entangled: if the checks-side filter is also backwards, fixing only the deposits side produces a half-corrected reconciliation, arguably worse than a consistently-wrong one. Decide both together.

**Option 2 — Treat it as a workflow gap, not a code defect.** If the "Uncl Deposit" checkbox is supposed to be used by staff and simply never has been, the fix is process (start flagging outstanding deposits), not code — but the popup will keep firing every single time until the first box is checked, which makes this impractical on its own without Option 3 alongside it.

**Option 3 — Narrow the guard so it warns on genuine anomalies only.** The guard currently treats "aggregate over zero matching rows" the same as "value couldn't be computed." Those are different conditions. If a permanently-empty `DepCleared` population is *accepted as legitimate* (pending Option 1), the guard could drop `IsNull(Me.txtSumOfUnclearDeposits)` from its condition — or, better, be split so an empty deposit population is reported distinctly (or silently treated as `$0.00`) while a `Null` `txtSumOfBalances` or `txtSumOfUncashed` still raises the alarm. This stops the popup without touching any reported figure, but it is a deliberate decision to stop warning about this specific case — do it only if Option 1's answer is "zero really is correct here," not as a way to quiet the box before that question is answered.

**What not to do:** don't blanket-`Nz()` the source query or delete the guard outright. That's the exact failure mode the original investigation warned against — a reconciliation that *looks* balanced because a missing figure was silently zeroed is worse than one that refuses to compute, in a system whose whole job is catching trust-account discrepancies.

## Patch — Option 3, guard narrowing — **APPLIED & VERIFIED 2026-09-02**

**Status: applied to the live Access app and confirmed working by the user** — the popup no longer appears from the Admin page's "Trust Account (Chron)" tile, the "Requery" button, or `frmTakeOffReconciliation`, and the reported totals were unchanged by the edit.

The corresponding extract files were then hand-updated to match (`frmTRUSTENTRIESCHRON.txt:2493-2498` and `:2528-2533`, `frmTakeOffReconciliation.txt:3330-3335`), each keeping the prior condition as a commented `LEGACY` line per repo convention. **These extract edits were made by hand, not by re-running the extractor** — if the text typed into Access differed from the patch below (e.g. the explanatory comment lines were skipped), a fresh export will supersede them.

**Decision:** stop warning on the `txtSumOfUnclearDeposits` branch only. No computed figure changes — `txt_WF_CHK` keeps the `Nz(..., 0)` that has always treated the empty `DepCleared` population as `$0.00`. The `DepCleared` semantics question (Option 1) stays open as separate, sign-off-gated work.

Written as a **single-line edit in each affected procedure** rather than a whole-procedure paste — the deployed code is now known exactly (see line references below), and a one-line change can't disturb the sentinel `DLookup` repoint or the `Nz()` arithmetic that must stay intact.

**Three call sites**, all confirmed in the `2026-09-02` extract:

| Form | Procedure | Extract line | Condition to edit |
|---|---|---|---|
| `frmTRUSTENTRIESCHRON` | `cmdRequery_Click` | `:2493` | inline `If IsNull(...) Then` |
| `frmTRUSTENTRIESCHRON` | `Form_Load` | `:2523` | inline `If IsNull(...) Then` |
| `frmTakeOffReconciliation` | `Form_Load` | `:3330` | `blnReconciliationHasNulls = ...` |

In the Access VBA editor (Alt+F11) → `Form_frmTRUSTENTRIESCHRON`, in **both** `Form_Load` **and** `cmdRequery_Click`, find this line:

```vba
    If IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits) Then
```

and replace it with (LEGACY line kept commented directly above, per repo convention):

```vba
    ' LEGACY (pre-2026-09-02 guard narrowing, see .docs/reconciliation-forms-null-calculation-followup-2026-09-02.md)
    'If IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits) Then
    ' txtSumOfUnclearDeposits deliberately excluded: [Trust Account].DepCleared has never been True for
    ' any row (0 of 44,799, TDate range 1899-12-30..2026-09-02), so qryReconciliation_sumOfUnclearedDeposits
    ' is permanently Null. The Nz() on the txt_WF_CHK line below treats it as $0.00 - which is the value
    ' txt_WF_CHK has always used. Warning on it fires every load/Requery without indicating a real problem.
    If IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Then
```

**Change nothing else — and in particular, do not delete the now-"unused"-looking assignment line.** `Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")` must stay exactly where it is: the control is still read by the `txt_WF_CHK` line below it, and removing the assignment would leave that control holding a stale value from the previous record/load rather than a fresh (if `Null`) one. Only the `If ... Then` condition changes. The `txt_WF_CHK` assignment itself, including its `Nz(Me.txtSumOfUnclearDeposits, 0)`, stays byte-for-byte as-is — that `Nz()` is what keeps the figure correct and unchanged.

### Verify

1. **Before editing anything**, open the form and **write down the Wells Fargo figure** (`txt_WF_CHK`) currently displayed — dismiss the popup to see it. This is the before/after invariant. (Don't compare against this doc's `761,604.20`; `txtSumOfBalances` is a live aggregate over `vwTakeOff_trust_account` and moves with every trust transaction — new `2026-09-02` rows were already landing while this was written. The number legitimately drifts; what must not change is *before-edit vs. after-edit*.)
2. **Debug → Compile** the project — must compile clean before proceeding.
3. Open the Admin page → **"Trust Account (Chron)"**. No popup should appear.
4. On the form, click **"Requery"**. No popup should appear (this is the second call site — if it still pops, only one of the two procedures was edited).
5. Confirm `txt_WF_CHK` **matches the figure recorded in step 1**. A changed figure means something beyond the guard condition was touched — most likely the assignment line or the `Nz()` — and should be backed out via the LEGACY comment.
6. Confirm both `txtSumOfBalances` and `txtSumOfUncashed` are still named in the edited condition, so the guard still fires for the branches that do indicate a real problem.
7. Recompile to `.accde` and redistribute per the normal process, then **re-export the objects to `database_assessment/TBCMS/extract/`** — the `2026-09-02` refresh is what made this diagnosis exact, so keep the habit.

### `frmTakeOffReconciliation` needs the same edit — confirmed, not conditional

The `2026-09-02` extract confirms this form has the guard deployed too (`frmTakeOffReconciliation.txt:3330`), with the identical `Or IsNull(Me.txtSumOfUnclearDeposits)` term, so it pops the same box (worded "Do not rely on **the totals on this form**") for the same permanent reason. It matters more than the Chron form: it's the one that runs the monthly `Insert Data` close.

Same treatment — narrow the condition, change nothing else:

```vba
    ' LEGACY (pre-2026-09-02 guard narrowing, see .docs/reconciliation-forms-null-calculation-followup-2026-09-02.md)
    'blnReconciliationHasNulls = IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed) Or IsNull(Me.txtSumOfUnclearDeposits)
    blnReconciliationHasNulls = IsNull(Me.txtSumOfBalances) Or IsNull(Me.txtSumOfUncashed)
```

Verify it the same way: note `txt_WF_CHK`, `txtTotalMinusInputBalance` and `txtWFTrustAndJMSMTrust` before the edit, and confirm all three are unchanged after. On this form the sentinel fix is also visibly working — `txtSMTrust` and `txtRLFDeeds` should read `$0.00`, not blank.

## Other next steps

1. ~~Confirm the firing branch via the Immediate window.~~ **No longer needed** — the `2026-09-02` extract plus the live `DepCleared` distribution settle it directly. (If you want the belt-and-braces check anyway, with `frmTrustEntriesChron` open: `? IsNull(Forms!frmTrustEntriesChron!txtSumOfBalances), IsNull(Forms!frmTrustEntriesChron!txtSumOfUncashed), IsNull(Forms!frmTrustEntriesChron!txtSumOfUnclearDeposits)` → expect `False False True`.)
2. ~~Check whether the sentinel-`CaseID` fix is deployed.~~ **Done and confirmed working** — see Finding 3. No action.
3. ~~Re-export the extract.~~ **Done, `2026-09-02T15:15:31Z`.** Keep doing it after the guard-narrowing edit lands.
4. **Still open — the `DepCleared` semantics question (Option 1 above).** The guard narrowing stops the popup without changing any figure, but it does not answer whether `txt_WF_CHK` *should* be subtracting `$0.00` for uncleared deposits. That needs accounting sign-off and should be decided together with the sibling `CheckCashed` / `txtSumOfUncashed` inversion the original investigation flagged and left open.
5. Everything in the original fix doc's "Recommended fix" section otherwise still stands.

## Primary sources

- Live queries, `awsql2022dev` / `TateByWater`, run 2026-09-02 (this session): `dbo.vwTakeOff_A` (CaseIDs 12664/23437/11946), `dbo.vwTakeOff_trust_account` (`SumOfBalance` aggregate), `dbo.[Trust Account]` (`CheckCashed=1` / `DepCleared=1` aggregates; `GROUP BY DepCleared`, `GROUP BY CheckCashed`, `GROUP BY Reconciled` distributions; `MIN`/`MAX(TDate)` and total row count).
- `database_assessment/TBCMS/extract/vba/forms/frmTRUSTENTRIESCHRON.txt:707` (`Label112`, caption `"Uncl Deposit"`, `Left=14415`) and `:1999-2000` (`DepCleared` checkbox, `ControlSource="DepCleared"`, `Left=14895`) — the label/field-semantics evidence; also `:559` (`"Uncashed\015\012Check"` caption) and `:1885-1886` (`CheckCashed` checkbox) for the sibling case.
- `database_assessment/TBCMS/extract/vba/forms/frmHomeAdmin.txt:5971-5972,6274-6275` — `cmdTrustAccountChron` button, opens `frmTrustEntriesChron`.
- **Refreshed extract, `app_manifest.json` `extractedAtUtc: 2026-09-02T15:15:31Z`** (source `TB CMS.SQL.accdb`) — the deployed state of record for every code citation in this doc:
  - `vba/forms/frmTRUSTENTRIESCHRON.txt:2485-2499` (`cmdRequery_Click`, guard at `:2493`) and `:2516-2529` (`Form_Load`, guard at `:2523`), with `LEGACY` blocks at `:2474-2483` / `:2506-2514`.
  - `vba/forms/frmTakeOffReconciliation.txt:3319-3347` (`Form_Load`), `blnReconciliationHasNulls` at `:3330`, sentinel lookups at `:3324-3325`.
  - `queries/qryTakeOffAvailBalance.sql` — the new unfiltered sentinel-lookup query (fix doc Step 1), confirming that fix shipped.
- `database_assessment/TBCMS/extract/queries/qryTakeOff.sql`, `qryTakeOff_A.sql`, `qryReconciliation_sumOfUnclearedDeposits.sql`, `qryTakeOff_uncleared_deposits.sql` — query definitions used to translate the live SQL checks faithfully.
- The superseded `2026-08-27` extract (`de31aa7`) — showed the pre-fix bodies, which is what made Finding 1's deployment state inferential on the first pass; retained here only to explain why that pass reasoned from the popup text rather than from source.
- [reconciliation-forms-null-calculation-investigation.md](reconciliation-forms-null-calculation-investigation.md) and [reconciliation-forms-null-calculation-fix.md](reconciliation-forms-null-calculation-fix.md), committed `83f72af` (2026-08-27 10:45:36 -0400) — prior investigation and proposed fix; Finding 1 is a direct textual comparison against the latter, Finding 3 restates and reconfirms that investigation's H2 against today's data.
- User-provided screenshots ("Admin" page and "Trust Account Entries" page, 2026-09-02) — the `$761,604.20` figure compared against Finding 2's computed value.
