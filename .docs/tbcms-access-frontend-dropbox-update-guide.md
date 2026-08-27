# TBCMS Access Frontend — Dropbox Integration Update Guide

**Audience:** an engineer sitting at MS Access with `TBCMS_Test.accde` (or an
equivalent test-linked `.accdb`) open, who needs to bring the frontend's VBA
up to date with the Dropbox/Bridge work that already exists as source in
`Dropbox-Migration/`.

**Scope:** import procedure only. This does not cover Phase 7 production
cutover, the bridge server's own deployment (see
`.docs/bridge-deployment-runbook.md`), or unresolved product decisions (D1,
DocumentType folder mapping, etc. — see the master plan's `Known Gaps`
section).

---

## Summary

The `database_assessment/TBCMS/extract/` snapshot was regenerated on
2026-08-27 (commit `de31aa7`) by static extraction (`SaveAsText` + Design
View reads, no code executed — confirmed by the stage list in
`extract/run_summary.json`) from a file the extractor recorded as
`C:\Users\indra\Downloads\TBCMS\TB CMS.SQL.accdb`
(`extract/app_manifest.json:2-3,17`). Its `DocumentManagement` module is 839
lines (`extract/vba/modules/DocumentManagement.txt`), contains zero mention
of Dropbox, and the module inventory captured in
`extract/app_manifest.json:3691-3761` lists 23 standard modules with **no**
`DropboxService` module and **no** class modules
(`extract/app_manifest.json:3762`). The frontend this snapshot represents has
never had the Dropbox work imported.

Two things are true at once and this guide is written around both:

1. **The captured frontend is behind** `Dropbox-Migration/` — as CLAUDE.md
   says, and as re-verified here.
2. **`Dropbox-Migration/` has itself moved past what CLAUDE.md describes.**
   CLAUDE.md documents a direct VBA→Dropbox OAuth/DPAPI design (kill-switch,
   mandatory `Dropbox-API-Path-Root` header injected from VBA,
   `tblDropboxTokens`/`tblDropboxLog` as required local tables). That design
   is still in the file, but it is now the **rollback path**, gated behind a
   module-level compiler directive `#Const PREBRIDGE_LEGACY = False`
   (`Dropbox-Migration/DropboxService.bas:115`). The **active** code path
   (committed through `e6f66ac`, 2026-06-29) routes every Dropbox call
   through a new internal REST service, `TBCMSDropboxBridge`
   (`dropbox-bridge/`, tracked in this repo), over Windows Integrated Auth —
   see the `BRIDGE REWRITE` header block at
   `Dropbox-Migration/DropboxService.bas:8-31`. `DropboxService.bas` is 4,480
   lines today (not the ~3,400 CLAUDE.md/the plan cite),
   `DocumentManagement.bas` is 2,092 lines (not ~1,900-1,953), and
   `Dropbox-Migration-SQL-Install.sql` is 2,332 lines across **9** sections
   (not 2,251 / 8) — Section 9 adds the two SQL objects the bridge needs
   (`Dropbox-Migration-SQL-Install.sql:2244-2331`). This is elaborated in
   **Discrepancies** below; the practical effect is that the "VBA
   prerequisites" for this import are different from what CLAUDE.md implies.

---

## Prerequisites

### 1. Confirm which file you are actually editing

The extract's source file (`TB CMS.SQL.accdb`, from `...\Downloads\TBCMS\`)
is **linked to production SQL Server**, not the test database: every
`ODBC;DRIVER=SQL Server;SERVER=tbf-cms;...` connect string in
`extract/schema.json` (e.g. lines 304, 408, 504, 640, 696, 1208, 1328, 1808,
1904, 2096) targets `tbf-cms`. That file was pulled down purely as a
point-in-time analysis snapshot (module/form/schema inventory) — it is
**not** `TBCMS_Test.accde` and should not be used as the import target as-is.

- If you're working in the actual `TBCMS_Test.accde` (or an `.accdb` you
  already know is linked to `awsql2022dev`), skip to step 2.
- If you are instead opening a copy of this same downloaded file: relink
  every ODBC-linked table from `tbf-cms` to `awsql2022dev` first (External
  Data → Linked Table Manager, or a `TableDef.Connect` + `RefreshLink` VBA
  script) and confirm the relink before doing anything else. Do **not** open
  it with `ALLOW_DROPBOX_WRITES` ever flipped `True` while it might still be
  pointed at `tbf-cms`. (Note: the app's own `Relinking` standard module,
  `extract/vba/modules/Relinking.txt`, relinks a *local* `Login_BE1.accdb`
  back end, not these SQL Server ODBC links — it will not do this for you.)

### 2. SQL Server side — verify, don't re-run blindly

The installer (`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql`) must
already be applied to the target database before the imported VBA will do
anything useful — verified live against `awsql2022dev/TateByWater` today via
`sqlcmd`:

| Check | Result |
|---|---|
| `tblDropboxConfig.BridgeUrl` column exists | Yes |
| `dbo.tblDropboxServiceToken` exists | Yes, 1 row |
| `tblDropboxConfig.BridgeUrl` current value | `http://localhost:8088/api` |
| `tblDropboxServiceToken` account | `sugianto@tatebywater.com`, `SetupByUser = auto-refresh`, last updated 2026-06-29 |
| `tblDropboxRootConfig` rows | 1 |
| `tblDropboxAuditLog` rows | 33 |

So Section 9 (`Dropbox-Migration-SQL-Install.sql:2244-2331`) **is** applied
on the dev DB, even though `.docs/dropbox-bridge-plan.md:20-22` still says
"Not yet applied to the dev DB" — that status line is stale, see
Discrepancies. Because `BridgeUrl` resolves to `localhost:8088`, the bridge
service is expected to run **on the same machine as the Access session**,
not on a shared dev server.

**Do not re-run the full installer to "make sure."** Its header warns this
is destructive (`Dropbox-Migration-SQL-Install.sql:55`): Section 1
`DROP`/`CREATE`s the 6 core Dropbox tables
(`Dropbox-Migration-SQL-Install.sql:98-103`), which wipes
`tblDropboxAuditLog`/`tblDropboxOrphanQueue`/`tblDropboxVerificationReport`
history and resets `AppSecret` to a placeholder. Section 9.1 then re-adds
`BridgeUrl` but reseeds it to the **placeholder** URL
(`Dropbox-Migration-SQL-Install.sql:2280`, `http://tbcms-bridge.tatebywater.local/api`)
— if you do re-run it, you must manually restore `BridgeUrl` to the real
running instance's URL afterward, exactly as for `AppSecret`
(see the installer's own closing reminder at
`Dropbox-Migration-SQL-Install.sql:2328-2329`).

### 3. A bridge instance must be reachable at the configured `BridgeUrl`

Nothing Dropbox-related will work — not even read-only smoke tests — unless
something is listening at `tblDropboxConfig.BridgeUrl`. For local
dev/test work that means running the bridge project yourself:

```
cd dropbox-bridge
dotnet run           # or the "http" launch profile — listens on http://localhost:8088
```

(`dropbox-bridge/Properties/launchSettings.json`, `http` profile,
`applicationUrl: "http://localhost:8088"` — matches the `BridgeUrl` value
found on `awsql2022dev` above.) `? DropboxService.PhaseBridge_ConnectivityTest`
(after import) is the smoke test that confirms this end to end.

**Open item to confirm with IT/the author before doing any write testing:**
the committed `dropbox-bridge/appsettings.json:16` currently has
`"Bridge": { "AllowWrites": true, ... }`. Git history shows this value was
introduced as `false` (commit `01cf2daf`), the most recent substantive
bridge commit (`e6f66ac`, 2026-06-29 14:26) states in its message *"Bridge
config (appsettings.json AllowWrites, launchSettings.json port) left
uncommitted on purpose — committed default must keep writes off,"* and then
8 minutes later, the very next commit (`53ff7d4`, "setting update",
2026-06-29 14:34) flipped exactly those two settings and **committed** them,
landing on `AllowWrites: true` as of `HEAD`. This may be an intentional
personal-dev-loop convenience left in by mistake, or a deliberate decision
that superseded the earlier note — the history doesn't say which. It is a
**server-side** gate, independent of the VBA kill switch below; with a live
service token already provisioned (`tblDropboxServiceToken`, above) a bridge
started from this committed config **can** write to the real `/Company`
tree if the VBA-side guard is ever bypassed or flipped. Confirm intent
before relying on this file's committed defaults for anything beyond
read-only testing.

### 4. Kill-switch state

`Public Const ALLOW_DROPBOX_WRITES As Boolean = False` —
`Dropbox-Migration/DropboxService.bas:124` (not "~line 86" as the plan's
NEXT-SESSION note says — see Discrepancies). Leave it `False` for this
import. Every gated write entrypoint calls `GuardWritesEnabled`
(`Dropbox-Migration/DropboxService.bas:583-609`) as its first statement,
independent of `#Const PREBRIDGE_LEGACY`, so this remains the primary
safety boundary on the VBA side in both the bridge and pre-bridge code
paths.

### 5. Know the two independent rollback mechanisms

- **`DocumentManagement.bas`** — comment-block rollback. Every rewired
  function is preceded by a `LEGACY (pre-Phase 4a/4b/4d)` commented-out
  block (exact markers at
  `Dropbox-Migration/DocumentManagement.bas:850,898,1203,1284,1525`, plus
  `205` and `285` for the two config functions). To roll back one function:
  comment the active version, uncomment its LEGACY block, re-import.
- **`DropboxService.bas`** — compiler-directive rollback. `#Const
  PREBRIDGE_LEGACY = False` at line 115 gates roughly two dozen `#If
  PREBRIDGE_LEGACY ... #End If` blocks throughout the file (OAuth,
  DPAPI, the old direct-Dropbox `HttpRequest` transport and its namespace-
  header injection, the `tblDropboxTokens` schema upgrade, the old Phase 3b
  smoke tests). Flipping the single constant to `True` and recompiling
  restores the entire pre-bridge module at once — there is no per-function
  swap here.

---

## Step-by-step import procedure

1. **Back up the target file** before touching it (copy the `.accdb`, or
   note the `.accde`'s current compiled state / have a known-good source to
   revert to).
2. **Open the VBE** (Alt+F11) in the target file.
3. **Remove the stale module(s) first, then re-import both `.bas` files
   together.** This order matters — the migration plan documents a real
   failure mode from getting it wrong (`.docs/dropbox-migration-plan.md:54`,
   "the re-import gotcha"): a stale or half-updated module fails to
   compile, which makes *all* its public members invisible to form callers,
   surfacing as a misleading `"...can't be found"` error rather than a
   compile error. The plan is explicit that re-importing only one of the
   two modules is not enough, because they reference each other — if the
   target already has a `DropboxService` module from a prior partial import
   attempt, remove that one too before importing either, so you never end
   up with one fresh module paired against one stale sibling.
   - Right-click `DocumentManagement` in the Project Explorer → **Remove
     DocumentManagement...** → decline the export prompt (the current
     source is already tracked in this repo/extract). If a `DropboxService`
     module already exists (from an earlier attempt), remove it too.
   - **File → Import File...** → `Dropbox-Migration/DocumentManagement.bas`.
     Because the file carries `Attribute VB_Name = "DocumentManagement"`
     (`Dropbox-Migration/DocumentManagement.bas:1`) and no module of that
     name exists anymore, it imports cleanly as `DocumentManagement` (an
     import performed *without* removing the old module first creates a
     `DocumentManagement1` duplicate instead).
   - **File → Import File...** → `Dropbox-Migration/DropboxService.bas`
     (brand new module — nothing to remove first).
   - Delete any stray `...1`-suffixed duplicate modules left over from a
     prior partial import attempt.
4. **Add the `frmHome` startup/shutdown wiring — this is a net-new form
   edit, not covered by either `.bas` import.** The migration plan and the
   `DropboxService.bas` header both describe `StartupBootstrap` /
   `StartupShutdown` as "wired into `frmHome.Form_Open` / `Form_Unload`"
   (`Dropbox-Migration/DropboxService.bas:82,3926,4034`), but the currently
   captured `frmHome` has **only** a `Form_Load` handler and no `Form_Open`
   or `Form_Unload` at all (`extract/vba/forms/frmHome.txt:2044-2054`).
   There is no tracked source for this wiring anywhere in the repo — it
   must be added by hand:
   - Open `frmHome` in Design view. On the property sheet, set **On Open**
     and **On Unload** to `[Event Procedure]` (if either is already set to
     `[Event Procedure]` with an existing empty stub, add the call inside
     that existing handler instead of creating a new one), then in the
     code-behind add:
     ```vb
     Private Sub Form_Open(Cancel As Integer)
         DropboxService.StartupBootstrap
     End Sub

     Private Sub Form_Unload(Cancel As Integer)
         DropboxService.StartupShutdown
     End Sub
     ```
   - (`StartupBootstrap` loads config / pings the bridge and surfaces a
     user-facing message on failure; `StartupShutdown` purges
     `%TEMP%\TBCMS\` — see `Dropbox-Migration/DropboxService.bas:4100-4105`
     and the block above it.) Note the app's `StartUpForm` is `frmLogin`
     (`extract/app_manifest.json:9`), so this only fires after a successful
     login opens `frmHome`, not at database open.
5. **VBA project references — nothing new to add, but confirm live.**
   Static extraction has no visibility into Tools ▸ References at all — the
   extraction pipeline's own stage list
   (`extract/run_summary.json:5-70`) has no "references"/"libraries" stage,
   and `extract/app_manifest.json` has no such key. Everything below is
   inferred from source text, not observed in the VBE, and should be spot-
   checked once the file is open live:
   - **ADODB (early-bound `New ADODB.Connection`/`.Recordset`/`.Command`)**
     — already required by the *existing* legacy `DocumentManagement.txt`
     (e.g. `extract/vba/modules/DocumentManagement.txt:15`), and used the
     same way by both new modules
     (`Dropbox-Migration/DocumentManagement.bas:85` etc.,
     `Dropbox-Migration/DropboxService.bas:389,396,470,475,725,730,...`).
     No new reference needed.
   - **DAO (early-bound `Dim db As DAO.Database`, `dbFailOnError`,
     `dbOpenDynaset`)** — also already in use elsewhere in the existing app
     (`extract/vba/modules/Configuration.txt:16-20,50-54`,
     `extract/vba/modules/Authentication.txt:27`), and used the same way in
     `DropboxService.bas`'s local-log and (legacy-only) token-table code
     (e.g. `Dropbox-Migration/DropboxService.bas:783-813, 905-919`). No new
     reference needed.
   - **WinHTTP, MSXML, `Scripting.FileSystemObject`, `ADODB.Stream`** — all
     accessed via late-bound `CreateObject(...)`
     (`Dropbox-Migration/DropboxService.bas:691,700,1407,1585,1632,1670,
     2502,2530,3019,3036,3064,3099,3789,3838,4366`) — no project reference
     required at all, by design.
   - **`CoCreateGuid`/`StringFromGUID2`/`CredUIPromptForCredentials`/
     `CryptProtectData`/`CryptUnprotectData`** — Win32 API `Declare
     PtrSafe`/`Declare` pairs against system DLLs (`ole32.dll`,
     `credui.dll`, `Crypt32.dll`, `kernel32`;
     `Dropbox-Migration/DropboxService.bas:247-368`), gated per-bitness by
     `#If VBA7`. These are not VBE references either; they just need the
     DLLs present (they ship with Windows).
6. **`tblDropboxLog` — no manual step needed.** `LogLocal` lazily
   self-creates it via `EnsureLocalLogTable`
   (`Dropbox-Migration/DropboxService.bas:777-815`), a plain DAO `CREATE
   TABLE`, on first call. It does not exist in the current extract
   (confirmed: no `tblDropboxLog`/`tblDropboxTokens` under
   `extract/tables/`) and doesn't need to.
7. **`tblDropboxTokens` — not needed under the current default build.**
   Its schema-upgrade routine, `UpgradeTokenTableSchema`, is entirely
   inside `#If PREBRIDGE_LEGACY Then`
   (`Dropbox-Migration/DropboxService.bas:900-...`) — dead code with
   `PREBRIDGE_LEGACY = False`. Only create/upgrade it if you deliberately
   flip `PREBRIDGE_LEGACY` to roll back to the pre-bridge direct-OAuth path
   (item 5 above), in which case that existing routine handles it, not a
   manual `CREATE TABLE`.
8. **Debug → Compile** (VBE menu) before running anything. Fix any
   reference/compile errors here — see item 5's caveats.
9. **Confirm state before testing:**
   - `ALLOW_DROPBOX_WRITES = False` — `Dropbox-Migration/DropboxService.bas:124`.
   - `#Const PREBRIDGE_LEGACY = False` — `Dropbox-Migration/DropboxService.bas:115`
     (leave `False` unless deliberately testing the rollback path).
   - A local bridge instance is running and `tblDropboxConfig.BridgeUrl`
     points at it (Prerequisites §3).
10. **Run the smoke tests** (Immediate window), in this order — this is the
    authoritative list read directly from `Dropbox-Migration/DropboxService.bas`
    (the plan's NEXT-SESSION smoke-test list,
    `.docs/dropbox-migration-plan.md:119-128`, predates the bridge rewrite:
    it names pre-bridge `Phase3b_*` tests that are now retired behind
    `PREBRIDGE_LEGACY` and would not compile/run as it describes them, and
    it does not mention the bridge-era test below at all):
    - `? DropboxService.PhaseBridge_ConnectivityTest` — added by the bridge
      rewrite (`.docs/dropbox-bridge-plan.md:34`), not in the plan's list.
      Confirms config loads (`BridgeUrl` present), the bridge answers over
      Windows Integrated Auth, `/api/status` reports `ok`
      (`Dropbox-Migration/DropboxService.bas:4157`).
    - `? DropboxService.Phase3a_SmokeTest` (`:854`)
    - `? DropboxService.Phase3c_SmokeTest` (`:2921`) — read-only ops.
    - `? DropboxService.Phase3d_SmokeTest` (`:3679`) — confirms every write
      entrypoint's guard fires.
    - `? DropboxService.Phase3e_SmokeTest` (`:4113`/`:4139`, `#If`/`#Else`
      pair — bridge-era version pings the bridge instead of checking a
      loaded OAuth token).
    - `? DropboxService.Phase4b_SmokeTest` (`:4425`) — expected result is
      `SKIP` (desktop-client routing retired).
    - Write-path re-verification (requires flipping
      `ALLOW_DROPBOX_WRITES` to `True`, re-import, run, flip back, re-import
      — do this deliberately, not by default):
      `? DropboxService.Phase4d_UploadSmokeTest` (`:3749`),
      `? DropboxService.Phase4e_CreateFolderSmokeTest` (`:3855`),
      `? DocumentManagement.Phase5_E2E_HappyPathTest(<testCaseID>)`
      (`Dropbox-Migration/DocumentManagement.bas:1759`; the plan used
      `CaseID = 30405` and enforces a pre-flight of zero existing
      `tblCaseDocuments` rows for that case).

---

## Forms to regression-test after import

Verified by grepping every form's `vbaCodeBehind` in
`database_assessment/TBCMS/extract/forms/*.json` for the `DocumentManagement`
public function names (word-boundary match, not requiring a literal call
syntax, to also catch bare-Sub-call style). **7 forms** have genuine calls —
see Discrepancies for why this is not the "8 forms" figure CLAUDE.md and
`.docs/dropbox-bridge-plan.md:97-98` both state.

| Form | Event handler(s) | Function(s) called |
|---|---|---|
| `frmClientLedger` | `cmdBillingOpenDocumentRetainer_Click`, `cmdCloseCase_Click`, `cmdCreateFolder_Click`, `cmdCreateFolderSub_Click`, `cmdOpenClosedFinal_Click`, `cmdOpenDocumentClientID_Click`, `cmdOpenDocumentFolderCorrespondence_Click`, `cmdOpenDocumentFolderFinance_Click`, `cmdOpenDocumentFolderFull_Click`, `cmdOpenDocumentFolderInvoices_Click`, `cmdOpenInitialIntake_Click`, `cmdOpenRetainer_Click`, `cmdReopenCase_Click`, `cmdScan_Click` | `OpenDocumentFile`, `OpenDocumentFolder`, `CopyDocumentToClosedFileScan`, `MoveDocumentByCaseStatus`, `GetScannerFolder`, `SelectFileDialog`, `SaveScannedFileAs` |
| `Intakes` | `cmdScan_Click` | `GetScannerFolder`, `SelectFileDialog`, `GetIntakeDocumentFileName`, `GetIntakeFolderName` |
| `Time_Keeping` | `cmdBillingOpenInvoiceFolder_Click`, `cmdRecordShortTK_Click`, `cmdRecordTKStatement_Click` | `OpenDocumentFolder`, `GetAllInvoicesFolderName`, `GetDocumentFolderName`, `SaveCaseDocument` |
| `frmInvoiceSent` | `cmdBillingOpenInvoiceFolder_Click`, `cmdRecondSentPDInvoice_Click`, `cmdRecordSentInvoice_Click` | `OpenDocumentFolder`, `GetClosedDocumentFolderName`, `GetDocumentFolderName`, `GetAllInvoicesFolderName`, `SaveCaseDocument` |
| `frm_invoices_summary` | `cmdRecordInvoice_Click`, `cmdRecordPDInvoice_Click` | `GetClosedDocumentFolderName`, `GetDocumentFolderName`, `GetAllInvoicesFolderName`, `SaveCaseDocument` |
| `frmTimeKeepingClosed` | `cmdRecordShort_Click`, `cmdRecord_Click` | `GetDocumentFolderName`, `GetAllInvoicesFolderName`, `SaveCaseDocument` |
| `frmPersInjProvider` | `cmdOpenDocumentFolderMedDocs_Click`, `cmdMedDocsfolder_Click` | `OpenDocumentFolder` |

Also regression-test: **`frmClientLedger.cmdMerge_Click`**
(`extract/vba/forms/frmClientLedger.txt:18622`) — hardcodes
`GetObject("S:\Merge docs\EOA - JDR - Master.docx", ...)`. It does **not**
call any `DocumentManagement` function and is untouched by this import, but
it is a live `S:\` dependency the migration plan's own status table flags
as unresolved ("Pending: mail-merge inventory ... no callers identified
yet," `.docs/dropbox-migration-plan.md:151`) — this *is* that caller, and
it will keep failing/pointing at the old file share regardless of this
import.

---

## Discrepancies / open questions found during research

1. **CLAUDE.md's architecture description predates the bridge rewrite.**
   CLAUDE.md describes VBA-direct OAuth/DPAPI, a mandatory VBA-side
   `Dropbox-API-Path-Root` header, and `tblDropboxTokens`/`tblDropboxLog` as
   load-bearing local tables. All of that is real but is now the
   `#Const PREBRIDGE_LEGACY = True` rollback path
   (`Dropbox-Migration/DropboxService.bas:115` and the ~24 `#If
   PREBRIDGE_LEGACY` blocks it gates). In the active default build, the
   namespace header is injected server-side by the bridge
   (`dropbox-bridge/Services/DropboxApiClient.cs:225-234,258`), not by VBA
   — the VBA-side injector (`DropboxPathRootHeader`,
   `Dropbox-Migration/DropboxService.bas:600-607`) and its only call sites
   (`:1416, 2533, 3102`) are all inside `#If PREBRIDGE_LEGACY` blocks and do
   not compile in the current default build.

2. **7 forms call `DocumentManagement`, not 8 — verified, not a guess.**
   CLAUDE.md and `.docs/dropbox-bridge-plan.md:97-98` both say 8. A
   literal-text grep for the function names across all
   `extract/forms/*.json` initially matched 8 files, but the 8th,
   `frmPersonalInjury.json`, matches only because a button on that form
   happens to be *named* `cmdOpenDocumentFolderFinance`
   (`extract/forms/frmPersonalInjury.json:1808`) — its property block has
   no click-event wiring, its full `vbaCodeBehind` (checked in full) has no
   `..._Click` handler for it, and the app has zero Access macros at all
   (`extract/app_manifest.json:3763`, `"macros": []`), so there is no
   non-VBA path for that control to reach any VBA function either. A
   control with no code-behind `_Click` sub and no macro subsystem to fall
   back on cannot call anything. Net: 7 forms have real calls (table
   above).

3. **Line-count / section-count drift across `Dropbox-Migration/` and its
   own docs.** `DropboxService.bas` is 4,480 lines (docs say ~3,400-3,561);
   `DocumentManagement.bas` is 2,092 lines (docs say ~1,900-1,953);
   `Dropbox-Migration-SQL-Install.sql` is 2,332 lines / 9 sections (docs
   say 2,251 / 8). `.docs/dropbox-migration-plan.md`'s
   `▶ NEXT SESSION: START HERE` block still reads "paused 2026-06-04," but
   `git log` shows five more commits after that
   (`da7539d, ddde628, ce962ac, 616020a, cd766ca, 2e85a1f` — the bridge
   plan itself — plus `0bce1af, e44cb17, e6f66ac` after that), the latest
   dated 2026-06-29. Treat the plan's status snapshot and NEXT-SESSION
   block as stale for anything bridge-related; `.docs/dropbox-bridge-plan.md`'s
   own `▶ Implementation status (2026-06-22)` block is the more current
   status doc, but even it is now behind — it says Section 9 is "not yet
   applied to the dev DB" (`.docs/dropbox-bridge-plan.md:20-22`), which this
   session's live `sqlcmd` check against `awsql2022dev` shows is no longer
   true (Prerequisites §2), and it says `DocumentManagement.bas` is
   "UNCHANGED" by the bridge work (`.docs/dropbox-bridge-plan.md:121`),
   which the later `e6f66ac` commit (Phase 4e UX fixes) contradicts.

4. **`dropbox-bridge/appsettings.json`'s committed `AllowWrites` value.**
   See Prerequisites §3 — currently `true` on `HEAD`, with git history
   showing a same-day flip from `false` shortly after a commit message
   stating the opposite should stay uncommitted. Flagged as an open
   question for IT/the author, not asserted as a bug.

5. **`ALLOW_DROPBOX_WRITES`'s line number moved.**
   `.docs/dropbox-migration-plan.md:69` says "~line 86"; it is at line 124
   in the current file. Minor, but worth using the live grep result rather
   than the plan's citation when scripting anything against it.

6. **`extract/vba/index.json`'s static call graph undercounts.** Its
   per-procedure `"calls"` field found exactly one match for any
   `DocumentManagement` function name across the *entire* 1,201-procedure
   inventory (`DocumentManagement` calling its own `OpenFileDialog`
   helper). The reliable method used for the forms table above was a
   direct text search over each form's raw `vbaCodeBehind`, not this
   index's `calls` field — don't rely on `index.json`'s call graph alone
   for this kind of cross-module question.

7. **Cannot verify from static extraction alone (confirm live in
   Access/VBE):**
   - The actual checked state of Tools ▸ References in the target
     `.accde`/`.accdb` — no extraction stage captures this
     (`extract/run_summary.json`'s stage list has no references/libraries
     stage). Section 5 above is inference from source code patterns
     (early- vs late-binding), not an observed reference list.
   - Whether the exact library versions already in use (e.g. "Microsoft DAO
     3.6 Object Library" vs a newer DAO version string) match what a fresh
     Access installation would offer — version string isn't recoverable
     from `SaveAsText` exports.
   - Whether `TBCMS_Test.accde`'s linked-table `Connect` strings actually
     point at `awsql2022dev` today — only the separately-downloaded `TB
     CMS.SQL.accdb` extract source was inspected here (§1), and it points
     at `tbf-cms`. Confirm the real test build's linkage directly.

---

## Sources

- `database_assessment/TBCMS/extract/app_manifest.json:2-3,9,17,3691-3762`
- `database_assessment/TBCMS/extract/run_summary.json:1-70`
- `database_assessment/TBCMS/extract/schema.json:304-305,408-409,504-505,640-641,696-697,1208-1209,1328-1329,1808-1809,1904-1905,2096-2097`
- `database_assessment/TBCMS/extract/vba/index.json` (1,201-procedure inventory; `calls`/`redFlags` fields)
- `database_assessment/TBCMS/extract/vba/modules/DocumentManagement.txt` (839 lines; function list at lines 7,45,83,118,156,192,228,266,335,374,417,487,519,552,588,613,705,769,806)
- `database_assessment/TBCMS/extract/vba/modules/Configuration.txt:16-20,50-54`
- `database_assessment/TBCMS/extract/vba/modules/Authentication.txt:27`
- `database_assessment/TBCMS/extract/vba/modules/Relinking.txt:1-30`
- `database_assessment/TBCMS/extract/vba/forms/frmHome.txt:2044-2054`
- `database_assessment/TBCMS/extract/vba/forms/frmClientLedger.txt:18622`
- `database_assessment/TBCMS/extract/forms/frmPersonalInjury.json:1806-1823` (and full `vbaCodeBehind`)
- `database_assessment/TBCMS/extract/forms/{frmClientLedger,Intakes,Time_Keeping,frmInvoiceSent,frm_invoices_summary,frmTimeKeepingClosed,frmPersInjProvider}.json` (`vbaCodeBehind` fields)
- `database_assessment/TBCMS/extract/tables/tblDocumentRootDirectory.md` (legacy table still present; no `tblDropbox*` tables present)
- `Dropbox-Migration/DropboxService.bas:1,4-32,34-100,115,124,175,246-368,382-520,583-609,600-607,777-919,1391-1433,1394-1433,1416,1565,1635,1673,2518-2555,2533,3087-3119,3102,3679,3749,3855,3926,4034,4100-4157,4425`
- `Dropbox-Migration/DocumentManagement.bas:1,3,6-75,78,205,285,564-567,850,898,970,1203,1246,1284,1373,1525,1586,1652,1684,1759`
- `Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:55,87,332,375,418,481,956,1181,1411,1474,2244-2331,2267-2287,2280,2292-2319,2328-2329`
- `.docs/dropbox-migration-plan.md:16,18,54,69,71-135,138-156`
- `.docs/dropbox-bridge-plan.md:1-125,16-42,97-98,104-125`
- `.docs/bridge-deployment-runbook.md:1-42`
- `dropbox-bridge/appsettings.json:16`
- `dropbox-bridge/Properties/launchSettings.json`
- `dropbox-bridge/Services/DropboxApiClient.cs:12,225-234,258`
- `git log` on `Dropbox-Migration/DropboxService.bas`, `Dropbox-Migration/DocumentManagement.bas`, `.docs/dropbox-bridge-plan.md`, `dropbox-bridge/appsettings.json` (commits `da7539d, ddde628, ce962ac, 616020a, cd766ca, 2e85a1f, c225db3, ea945a3, 03b7bdd, 72db97f, 38b5b37, cc78413, 3744b91, 0fbc14f, 3cfbddb, 0bce1af, e44cb17, e6f66ac, 53ff7d4, 01cf2daf`); `git ls-files | wc -l` = 953 as of this session (CLAUDE.md's "17 tracked files" figure no longer holds)
- Live `sqlcmd -S awsql2022dev -d TateByWater` queries against `dbo.tblDropboxConfig`, `dbo.tblDropboxServiceToken`, `dbo.tblDropboxRootConfig`, `dbo.tblDropboxAuditLog`, run during this research session (2026-08-27)
