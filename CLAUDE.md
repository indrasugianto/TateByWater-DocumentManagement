# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository actually is

This is **not** the Python "database assessment tool" that `README.md` describes. That tooling (`assess_access_db.py`, `extract_vba.py`) was removed; `README.md` and the eight `.cursor/rules/*.mdc` files are **stale** — they document the old Python project and Python coding standards that no longer apply. Treat them as historical unless a task is explicitly about them.

The live work is a **migration of TBCMS (Tate By Water Case Management System) — a law-firm document-management subsystem — from on-prem `S:\` file shares to Dropbox Business via the Dropbox API.** TBCMS is a split MS Access (VBA) frontend + SQL Server backend application. This repo holds the migration's source-of-truth artifacts and planning docs; it is not a runnable application itself.

## The working set (everything else is reference)

| Path | What it is |
|---|---|
| [Dropbox-Migration/DropboxService.bas](Dropbox-Migration/DropboxService.bas) | ~3400-line VBA module: the entire Dropbox API integration (OAuth, tokens, read/write ops, audit logging). New module, built for this migration. |
| [Dropbox-Migration/DocumentManagement.bas](Dropbox-Migration/DocumentManagement.bas) | ~1900-line VBA module: the legacy document-operations module, rewired to delegate to `DropboxService` instead of `S:\` filesystem calls. Public function signatures are frozen (8 forms call them). |
| [Dropbox-Migration/Dropbox-Migration-SQL-Install.sql](Dropbox-Migration/Dropbox-Migration-SQL-Install.sql) | 2251-line idempotent SQL installer: schema, data-quality fixes, the one-time `S:\`→`/Company/` path migration, and the stored-procedure rewrites. |
| [.docs/dropbox-migration-plan.md](.docs/dropbox-migration-plan.md) | **The master plan (~167 KB). Start every session at its `▶ NEXT SESSION: START HERE` block** — it holds live status, what changed last session, and the next concrete step. |
| [.docs/document-management-analysis.md](.docs/document-management-analysis.md) | Grounded current-state analysis of the legacy subsystem (storage roots, 29 document types, 11 stored procs, data-quality defects). Read before changing any path/SP logic. |

These `.bas` files are **edited as text here, then imported into the Access frontend** (`TBCMS_Test.accde`) to run. The `.accdb`/`.accde` files themselves are gitignored (binary). There is no build/compile step in this repo — compiling the `.accde` happens inside MS Access and is documented in the (pending) IT runbook.

## Architecture you must hold in your head

- **Split app.** Frontend = per-user Access `.accde` holding all VBA + a few local tables (`tblDropboxTokens`, `tblDropboxLog`) accessed via DAO (`CurrentDb`). Backend = SQL Server holding all case data, document metadata, config, and **all path logic** in stored procedures, accessed via ADO using the existing helper `PcaGetConnnectionString()`.
- **VBA never builds a path.** Every folder/filename is resolved by a stored proc (`spGetDocumentFolderName`, `spGetDocumentFileName`, etc.) that tokenizes a naming template through `fnGetListOfWords`, substitutes columns from the view `vwfrmClientLedger`, and `EXEC`s dynamic SQL. The migration kept this engine and rerouted it from `tblDocumentRootDirectory` (S:\ roots) to `tblDropboxRootConfig` (`/Company/` roots), wrapping output in `REPLACE('\','/')`.
- **Two environments — never mix them.** This is the project's single largest risk.
  - **Test (Phases 0–6, where all current work happens):** SQL `awsql2022dev/TateByWater` (dev mirror) + `TBCMS_Test.accde`. `/Company` on Dropbox is treated **read-only**.
  - **Production (untouched until Phase 7 cutover):** still 100% on `S:\`, original `.accde`, SQL host `tbf-cms`. No script here runs against production until cutover.
- **The kill-switch is the safety boundary.** `Public Const ALLOW_DROPBOX_WRITES As Boolean = False` (top of `DropboxService.bas`). Every write entrypoint (`UploadFile`, `UploadLargeFile`, `MoveFile`, `CopyFile`, `DeleteFile`, `CreateFolder`) calls `GuardWritesEnabled` as its first statement and raises `vbObjectError + 6001` when writes are off. It is flipped to `True` only in the production build. To exercise a write test you flip it in source, run, then **flip it back** before committing.
- **Team namespace header is mandatory.** Every Dropbox API call must inject `Dropbox-API-Path-Root: {"namespace_id":"14334595683",...}` or it resolves against the user's empty personal namespace instead of the shared team tree.

## Conventions specific to this codebase

- **Rollback via preserved legacy blocks.** Every rewired VBA function and every rewritten stored proc is immediately preceded by a commented-out `LEGACY (pre-Phase 4x)` block containing the original body. To roll back one function: comment the active version, uncomment the LEGACY block, re-import/re-run. Preserve this pattern when you rewire anything.
- **"G-guards" (G1–G27)** are the numbered design decisions / known-gaps registry in the plan's `Known Gaps / Open Decisions` section. Code and commit messages reference them (e.g. "G2 `spMoveDocumentFolder` rewrite", "G13 token guard"). When you touch a guarded area, read that G-entry first and keep the reference in comments/commits.
- **Phased status lives in three places that must agree:** the plan's status snapshot + NEXT-SESSION block, and the file-header status comments at the top of each `.bas`. Update all of them when you change phase state.
- **Smoke tests are VBA functions, not a test runner.** There is no CI and no `pytest`/`xunit`. Validation = running functions like `? DropboxService.Phase3c_SmokeTest` in the Access VBA Immediate window, plus SQL verification queries. Rubberduck (rubberduckvba.com) is the named unit-test framework but harnessed tests are minimal. SQL changes are verified by the installer's Section 6 + by `sqlcmd`/SSMS queries.
- **Path drift in docs:** the plan and analysis sometimes reference the VBA extract at `msaccess/TBCMS/extract/...`; on disk the (largely untracked) extract is under `database_assessment/TBCMS/extract/`. Only 17 files are git-tracked (`git ls-files`); `database_assessment/`, `*.accdb`, `.claude/`, and `Dropbox-Migration/_temp/` are gitignored.

## Commands

There is no package manager, build, or lint step. The two things you actually run:

**SQL Server (test DB) — inspect, run verification, or apply the installer.** Windows/integrated auth works on this machine:
```powershell
# ad-hoc query
& sqlcmd -S awsql2022dev -d TateByWater -E -W -Q "SELECT COUNT(*) FROM tblCaseDocuments;"

# run the idempotent installer (see DESTRUCTIVE warning below before re-running)
& sqlcmd -S awsql2022dev -d TateByWater -E -i Dropbox-Migration\Dropbox-Migration-SQL-Install.sql
```
SQL-auth credentials (used by the Access frontend) live in `.claude/settings.local.json` (gitignored). Prefer `-E` here.

**⚠ The installer is DESTRUCTIVE on re-run.** Section 1 does `DROP ... CREATE` on the 6 Dropbox tables, wiping `tblDropboxAuditLog`, `tblDropboxOrphanQueue`, and `tblDropboxVerificationReport` history, and resets `AppSecret` to a placeholder (IT must re-`UPDATE` it after). It is idempotent for data fixes (WHERE-guarded) but not for that schema/history. Read its header comment before re-running on a non-fresh DB.

**VBA smoke tests** — run in the Access Immediate window after importing the module, e.g. `? DropboxService.Phase3d_SmokeTest` (confirms every write guard fires) or `? DocumentManagement.Phase5_E2E_HappyPathTest(30405)`.

**Git:** branch is `main`; commit messages use a `feat:`/`docs:`/`chore:`/`review:` prefix and `Phase Nx` tags matching the plan.

## SQL installer section map

Section 1 Phase 2 schema (6 Dropbox tables + `spLogDropboxAuditEvent`) · 2 manual-triage tables · 3 `tblDocumentTypes` typo fix · 4 intake natural-key fixes · 5 the `S:\`→`/Company/` path migration (11 per-category passes, auto-commits only if leftover-offender count = 0) · 6 verification · 7 Phase 1b diagnostic listings (output only) · 8 the stored-procedure rewrites (path-building SPs, G2 `spMoveDocumentFolder`, G13 `spSaveCaseDocument` guard) with legacy bodies preserved.

## Security note

Live credentials are present in tracked/untracked files in this repo: the Dropbox AppKey (`dqleswbnux8k3m5`) and namespace ID appear in the plan; SQL Server logins appear in `.claude/settings.local.json` and historically in `z_PCADataSources.csv`. The plan calls for rotating these (AppSecret + `TateBywaterSQLUser` in lockstep) before the production cutover. Do not propagate these secrets into new tracked files, and do not paste the SQL password into committed docs.
