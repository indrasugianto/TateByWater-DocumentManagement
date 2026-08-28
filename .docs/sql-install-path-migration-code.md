# Path-migration code in `Dropbox-Migration-SQL-Install.sql`

**Scope note:** primary source for this doc is `Dropbox-Migration/Dropbox-Migration-SQL-Install.sql` (2332 lines, read in full). `Dropbox-Migration/DropboxService.bas` and `Dropbox-Migration/DocumentManagement.bas` are cited only as secondary/cross-reference sources to confirm what the SQL-side path logic is for — they are called out explicitly wherever used. No source file was modified for this investigation.

## Summary

The installer touches directory/file paths in three distinct ways. First, **Section 5** does the one-time bulk data migration: it rewrites every stored `S:\...` path in `tblCaseDocuments.DocumentFileName`, `tblScans.ScanLocation`, and `[TB Intakes].[Scan Location GI]` to a `/Company/...` Dropbox path, across 11 pattern-specific `UPDATE` passes wrapped in one transaction, gated by a same-transaction leftover-offender check that rolls back and throws on any non-zero count. Second, **Section 8** rewrites the stored procedures that *build* paths at runtime (`spGetDocumentFolderName`, `spGetClosedDocumentFolderName`, `spGetClosedFileScanFolderName`, `spGetAllInvoicesFolderName`, `spGetIntakeFolderName`) so they resolve templates against the new `tblDropboxRootConfig` table instead of the legacy `tblDocumentRootDirectory`, plus the `spMoveDocumentFolder` (G2) and `spSaveCaseDocument` (G13) rewrites, all with legacy bodies preserved in commented-out blocks. Third, **Section 1** creates `tblDropboxRootConfig` itself — the config table whose seeded row (`/Company/COMMON` + four naming templates) is what Section 8's SPs read. Section 4's intake fixes and Section 6's verification are also examined below because CLAUDE.md's section map flagged them as possibly path-relevant; the finding is that Section 4 does **not** touch any path column, and Section 6 is read-only verification, not a second migration pass. A previously undocumented Section 9 (Dropbox Bridge service objects) exists in the file but carries no document-path logic.

---

## 1. `tblDropboxRootConfig` — the path-template config table

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:109-146`** (Section 1.1)

Creates the single-row config table that Section 8's path-building stored procedures read at runtime, and seeds it:

```sql
create table dbo.tblDropboxRootConfig
(
    ConfigID int not null primary key,
    NamespaceId varchar(50) not null,
    TeamRootPath varchar(500) not null,
    DocumentRootNaming varchar(500) not null,
    DocumentClosedNaming varchar(500) not null,
    ...
    constraint CK_DropboxRootConfig_SingleRow check (ConfigID = 1)
);
```
(`:109-123`)

```sql
values
(1, N'14334595683', N'/Company/COMMON',
 N'\ [Orig_Atty] \_CLIENTS\ [Case_Letter] \ [Last_Name] , ~ [First_Name] ~ [FileNo] \',
 N'\ [Orig_Atty] \_CLIENTS\ [Case_Letter] \ _CLOSED \ [Last_Name] , ~ [First_Name] ~ [FileNo] \',
 N'/Company/COMMON/_ALL INVOICES', N'', N'/Company/Closed File Scans', N'\ TB \ [Yr] \', N'/Company/COMMON/_SCANNER',
 N'/Company/COMMON/Intakes');
```
(`:140-145`)

This is the direct successor to the legacy `tblDocumentRootDirectory` table (referenced throughout the LEGACY blocks in Section 8, e.g. `:1481-1482`, `:1529-1530`). Notably the naming templates (`DocumentRootNaming`, `DocumentClosedNaming`, `ClosedFileScanNaming`) are still stored with **backslash** separators (`\ [Orig_Atty] \_CLIENTS\...`), unchanged from the legacy `S:\`-era format — the conversion to forward slashes happens downstream in each SP's output (see Section 8 below), not here. `AllInvoicesNaming` is seeded as an **empty string** (`N''` at `:144`), which is the reason `spGetAllInvoicesFolderName` (8.5) needs a distinct guard (see §8.5 below).

`NamespaceId` (`N'14334595683'`) is the Dropbox team-namespace ID CLAUDE.md's "Team namespace header is mandatory" note refers to; it is read by VBA and injected as the `Dropbox-API-Path-Root` header (cross-reference: `DropboxService.bas:607`, `:1416`, `:2533`, `:3102`).

Table is `DROP ... CREATE`d unconditionally on every run of the installer (`:103`, `:109`), which is why the DESTRUCTIVE warning at `:55-66` calls out that "seed values reset to plan-defaults" on re-run.

---

## 2. Section 2 — manual-triage tables (path-quarantine targets)

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:339-368`**

Creates (if missing — `object_id(...) is null` guarded, so these survive re-runs, unlike the Dropbox tables) the two tables Section 5 quarantines unrewritable path values into:

- `dbo.tblScans_ManualTriage` (`:341-349`) — `ScansID`, `OriginalValue` (the raw unrewritten `ScanLocation`), `Reason`, `QuarantinedAt`.
- `dbo.TBIntakes_ManualTriage` (`:355-366`) — `IntakeID`, `GILastName`/`GIFirstName`/`GIDate` (natural-key convenience columns), `OriginalValue` (the raw unrewritten `[Scan Location GI]`), `Reason`, `QuarantinedAt`.

These are the destination of the two quarantine `INSERT`s inside Section 5 (see B-7 and C-5 below) and the source of the Section 7.4/7.5 diagnostic listings.

---

## 3. Section 3 — `tblDocumentTypes` naming-rule typo fix

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:384-414`**

Not itself a path rewrite, but directly upstream of the path-*building* tokenizer (fixes a typo in `DocumentNamingRule`, the column `fnGetListOfWords` tokenizes when Section 8's SPs assemble a folder name):

```sql
update dbo.tblDocumentTypes
set DocumentNamingRule = replace(DocumentNamingRule, '(customeruserentry)', '(customuserentry)')
where DocumentTypeID = 30
      and DocumentNamingRule like '%(customeruserentry)%';
```
(`:387-390`)

Guard: wrapped in `begin try / begin transaction` (`:384-385`); after the `UPDATE`, it re-counts remaining offenders (`:395-398`) and `rollback`s + `throw`s error 51001 if any remain (`:400-404`); otherwise `commit` (`:406`). Standard `begin catch` rollback-and-rethrow at `:409-413`.

This typo (`(customeruserentry)` vs. the tokenizer-recognized `(customuserentry)`) is the same token class the new G13 guard in `spSaveCaseDocument` (§8.7 below) checks for at save time (`Dropbox-Migration-SQL-Install.sql:2180`).

---

## 4. Section 4 — intake natural-key fixes (does **not** touch path columns)

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:436-477`**

CLAUDE.md's section map flags this as possibly path-relevant ("intake natural-key fixes … may touch paths too"). Reading the actual `UPDATE` statements: **it does not.** All four updates write only `[GI First Name]` and/or `[GI Last Name]` on `dbo.[TB Intakes]`:

```sql
update dbo.[TB Intakes]
set [GI First Name] = N''
where [ID] = 183 and [GI First Name] is null;

update dbo.[TB Intakes]
set [GI Last Name] = N'Morales Bolanos', [GI First Name] = N'Marvin Eduardo'
where [ID] = 185 and ([GI Last Name] = N'Morales Bolanos, Marvin Eduardo' or [GI First Name] is null);

update dbo.[TB Intakes]
set [GI Last Name] = N'Carpio', [GI First Name] = N'Alejandro'
where [ID] = 193 and ([GI Last Name] is null or [GI First Name] is null);

update dbo.[TB Intakes]
set [GI Last Name] = N'Kerr'
where [ID] = 2482 and [GI Last Name] is null;
```
(`:439-467`)

The header comment (`:426-427`) notes that for two of these rows (193, 2482) the recovered name values were sourced *from* the row's `[Scan Location GI]` path — but the `[Scan Location GI]` column itself is never written by this section. The path column for these same rows is rewritten later, generically, by Section 5 Part C (§7 below) alongside every other `[TB Intakes]` row. Guard: single `begin try/transaction ... commit`/`begin catch ... rollback; throw` wrapper (`:436-437`, `:469-476`), no leftover-check gate (unlike Section 3 and Section 5) since these are fixed-ID, idempotent, WHERE-guarded single-row updates.

---

## 5. Section 5 — the `S:\` → `/Company/` path migration (the core of this task)

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:480-952`**

Header (`:480-490`) states the policy directly: rewrites three columns across "11 per-category passes," quarantines unrewritable rows, and **auto-commits only if the leftover-offender count = 0** across all three tables (excluding quarantined rows); otherwise `ROLLBACK + THROW`.

The whole section is one named transaction:

```sql
begin try
    begin transaction MigratePathsToDropbox;
    ...
end try
begin catch
    if xact_state() <> 0
        rollback transaction MigratePathsToDropbox;
    throw;
end catch;
```
(`:496-497`, `:947-951`)

### Part A — `tblCaseDocuments.DocumentFileName` (1 pass)

**`:499-535`**

Single `UPDATE` handles both a bare `S:\...` path and an Access-hyperlink-wrapped `#S:\...#`/`#S:\...` variant, then does the root swap and slash-direction swap:

```sql
update dbo.tblCaseDocuments
set DocumentFileName = replace(replace(   case
                                              when left(...)='#' and right(...)='#' then substring(...) -- strip both hashes
                                              when left(...)='#' then substring(...)                     -- strip leading hash
                                              when right(...)='#' then left(...)                         -- strip trailing hash
                                              else DocumentFileName
                                          end,
                                          'S:\', '/Company/'
                                      ),
                               '\', '/'
                              )
where DocumentFileName is not null and DocumentFileName <> ''
      and (DocumentFileName like 'S:\%' or DocumentFileName like '#S:\%');
```
(`:510-534`, condensed)

### Part B — `tblScans.ScanLocation` (6 rewrite passes + 1 quarantine pass)

**`:537-739`**, comment at `:538` explicitly says "six per-category passes + quarantine":

| Pass | Lines | Pattern matched | What it does |
|---|---|---|---|
| B-1 | `:554-580` | `'#S:\%'` | Access hyperlink wrapper around a bare `S:\` path — strips hash(es), decodes `%20`→space, then `S:\`→`/Company/`, `\`→`/` |
| B-2 | `:582-600` | `'S:\%'` | `displaytext#\\TBF-SRVR12\...#` hyperlink shape — truncates at the first `#` (keeps only the display text) before the same root/slash swap |
| B-3 | `:602-628` | `'#file:///S:\%'` | URL-encoded `file://` wrapper — strips the `#file:///` prefix (offset `2+8`) and hash suffix, decodes `%20`, then root/slash swap |
| B-4 | `:630-652` | `'#?S:\%'` | `#?` typo-prefixed hyperlink — strips the 2-char prefix and hash suffix, then root/slash swap |
| B-5 | `:654-684` | `'#file:///\\TBF-SRVR12\%'` | Legacy UNC path wrapped in `file://` — additionally rewrites the literal UNC root `\\TBF-SRVR12\Company\` → `/Company/` before the generic `S:\` swap |
| B-6 | `:686-712` | `'#\\TBF-SRVR12\%'` | Bare legacy-UNC hyperlink (no `file://`) — same UNC-root rewrite as B-5 |
| B-7 (quarantine) | `:714-739` | (none of the above, and not already `/Company/%`) | `INSERT`s the untouched row into `tblScans_ManualTriage` with reason `'Hash-less or mid-string-corrupted ScanLocation — no automatic rewrite'`, deduped via `not exists (... where q.ScansID = s.ScansID)` |

Sample of the UNC-root rewrite unique to B-5/B-6:
```sql
... replace(..., '\\TBF-SRVR12\Company\', '/Company/'), 'S:\', '/Company/'), '\', '/')
where ScanLocation like '#\\TBF-SRVR12\%';
```
(`:687-711`, condensed)

### Part C — `[TB Intakes].[Scan Location GI]` (4 rewrite passes + 1 quarantine pass)

**`:741-856`**, comment at `:742` says "four passes + quarantine":

| Pass | Lines | Pattern matched | What it does |
|---|---|---|---|
| C-1 | `:756-760` | `'S:\%'` | Bare path — `S:\`→`/Company/`, `\`→`/`. No `%20` decode step (unlike B-2/B-3/B-5 in Part B). |
| C-2 | `:762-784` | `'#S:\%'` | Hash-wrapped — strips hash(es), then root/slash swap |
| C-3 | `:786-812` | `'#file:///S:\%'` | `file://`-wrapped — strips `#file:///` prefix and hash suffix, decodes `%20`, then root/slash swap |
| C-4 | `:814-826` | `'?S:\%'` | `?`-prefixed — `substring(col, 2, len(col)-1)` strips the leading `?` but the length expression does **not** separately deduct a trailing `#` the way C-2/C-3 do, then root/slash swap |
| C-5 (quarantine) | `:828-856` | (none of the above, and not already `/Company/%`) | `INSERT`s into `TBIntakes_ManualTriage` (capturing `[ID]`, `[GI Last Name]`, `[GI First Name]`, `[GI Date]`, the raw `[Scan Location GI]`) with reason `'Hash-less or root-missing Scan Location GI — no automatic rewrite'`, deduped by `IntakeID` |

**11-pass count, reconciled:** Part A (1) + Part B rewrite passes B-1..B-6 (6) + Part C rewrite passes C-1..C-4 (4) = **11**, matching CLAUDE.md and the file's own section header (`:481-490`). The two quarantine `INSERT`s (B-7, C-5) are *not* counted among the 11 — they're labeled as extensions of Part B/Part C ("six per-category passes + **quarantine**", "four passes + **quarantine**") rather than as passes themselves. Labeling is also inconsistent between the two: Part B's quarantine step is called "B-7" (as if a 7th rewrite pass) while Part C's is "C-5" — same role, different numbering convention.

**Passes are narrower than the gate that validates them, by design.** C-1 does no `%20`-decoding (unlike B-2/B-3/B-5), and C-4's `substring(col, 2, len(col)-1)` strips the leading `?` but keeps any trailing `#` (unlike C-2/C-3, which explicitly deduct it via `case when right(...)='#'`). Whether this matters for any real row is a data question this script can't answer from its text alone — but structurally, it doesn't need each individual pass to be exhaustive, because the leftover-offender gate immediately below re-scans for exactly these residues (trailing `#` at `:915`, `%20` via the `[%]20%` pattern at `:918`) and rolls back the whole section if any survive. The gate, not the individual passes, is what makes Section 5 safe.

### The commit gate — leftover-offender check + auto-commit/rollback

**`:858-951`**

Inside the same transaction, after all passes, re-counts rows in each of the three tables that still look like an unrewritten/partially-rewritten path — containing a backslash, `S:` substring, leading/trailing `#`, `file:///`, the legacy UNC root, or a literal `%20` — while excluding rows already quarantined:

```sql
select @LeftoverDocs = count(*)
from dbo.tblCaseDocuments
where DocumentFileName is not null and DocumentFileName <> ''
      and ( DocumentFileName like '%\%'
         or DocumentFileName like '%S:%'
         or left(DocumentFileName,1)='#' or right(DocumentFileName,1)='#'
         or DocumentFileName like '%file:///%'
         or DocumentFileName like '%\\TBF-SRVR12\%'
         or DocumentFileName like '%[%]20%' collate Latin1_General_BIN );
```
(`:866-879`, one of three near-identical blocks — `tblScans` at `:881-898` additionally excludes quarantined `ScansID`s, `[TB Intakes]` at `:900-919` additionally excludes quarantined `OriginalValue`s)

Gate and outcome:
```sql
if @LeftoverDocs + @LeftoverScans + @LeftoverIntakes > 0
begin
    rollback transaction MigratePathsToDropbox;
    ... throw 51002, @msg, 1;
end;
commit transaction MigratePathsToDropbox;
```
(`:921-931`) — this is the "auto-commits only if leftover-offender count = 0" behavior CLAUDE.md describes. On success, per-pass row counts are printed (`:933-945`).

---

## 6. Section 6 — post-migration verification (read-only, not a second migration pass)

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:955-1177`**

Two blocks, both plain `SELECT`s outside any transaction — they report, they do not write:

- **`:964-1135`** — re-runs the same seven offender predicates from the Section 5 gate (backslash / `S:` / leading-hash / trailing-hash / `file:///` / legacy-UNC / `%20`) against all three tables post-commit, as a `sum(case when ... then 1 else 0 end)` column per predicate. Comment at `:957`: "Expected: every numeric column in the leftover-offender block reads 0."
- **`:1141-1177`** — path-prefix distribution: `left(<path-column>, 30)` grouped and counted, descending, for each of the three tables. Informational only.

This is standalone from, and runs after, the transaction described in Section 5 — it is verification of the committed result, not a second enforcement gate.

---

## 7. Section 7 — Phase 1b diagnostic listings (output only, path-relevant items)

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:1180-1373`**

All five sub-listings are `SELECT`-only (no `UPDATE`/`INSERT`); three are directly path-related:

- **7.1 — `:1188-1271`** (B-9): path-length pre-flight. Lists any post-migration path over 260 characters (Dropbox's effective path-length limit, tagged "G14" at `:1191`) in all three tables, plus a length-bucket histogram for `tblCaseDocuments`. Comment at `:1192-1193` records the test-DB baseline: zero rows over 260, max 247 chars.
- **7.2 — `:1274-1305`** (B-4): the 13 bracket-literal rows in `tblCaseDocuments` — i.e. rows where a template token like `[Case_Letter]` survived unresolved into the stored filename. Classifies each into "unresolved template (re-resolve via SPs)" vs. "truncated filename ending `[df`" vs. other, via `LIKE '%[[]Case[_]Letter]%'` / `LIKE '%[[]df'` (`:1287-1293`). The comment at `:1277-1280` is explicit that the fix path for the first class is **re-running `spGetDocumentFolderName` + `spGetDocumentFileName`** (Section 8 SPs) once the underlying `vwfrmClientLedger` row is repopulated — directly tying this listing to the path-building SPs in Section 8.
- **7.3 — `:1307-1344`** (B-6): non-canonical roots — rows whose (already-migrated) path does not match either canonical prefix (`/Company/COMMON/%/_CLIENTS/%` or `/Company/Closed File Scans/%`), grouped by first-60-char prefix, plus a 30-row sample. Because this listing filters on `/Company/%` (`:1323`), it necessarily inspects *post*-Section-5 values.
- **7.4 (`:1346-1357`, B-7) and 7.5 (`:1359-1373`, B-8)** simply dump the two manual-triage tables populated by Section 5's B-7/C-5 quarantine inserts, for IT hand-review.

Between Section 7 and Section 8 sits an unnumbered **"POST-INSTALL CONFIRMATION"** block (`:1376-1408`) that, among other sanity selects, reads back `tblDropboxRootConfig.TeamRootPath` (`:1393-1396`) — worth noting only because it's easy to miss: it isn't under any `SECTION n` banner.

---

## 8. Section 8 — path-building stored procedure rewrites (Phase 4c)

**`Dropbox-Migration/Dropbox-Migration-SQL-Install.sql:1410-2240`**

### Header and design notes (`:1410-1460`)

States the plan: 7 SPs listed (`:1414-1434`), 4 SPs deliberately **not** touched (`:1436-1444`):

> `spGetCaseDocument` — already returns `tblCaseDocuments.DocumentFileName` verbatim (now `/Company/`-rooted post-migration).
> `spGetDocumentFileName` — produces a FILENAME from `DocumentNamingRule` (no path separator concerns).
> `spGetIntakeDocumentFileName` — same, for intake filename.
> `fnGetListOfWords` — tokenizer used by the path-building SPs; preserved verbatim (works for both `/` and `\`).
(`:1438-1444`)

This directly answers the task's question about `spGetDocumentFileName` and `fnGetListOfWords`: **neither is rewritten**, and the file states why — filename-building has no directory-separator concerns, and the word-tokenizer is separator-agnostic so it needs no change regardless of which root format it's fed.

Path-separator design note (`:1451-1456`): `tblDropboxRootConfig` keeps templates in legacy backslash form; every rewritten SP instead wraps its *dynamic-SQL output* in `REPLACE(..., '\', '/')` so only the final resolved path is forward-slash. Rollback procedure stated at `:1446-1449`: comment out the new `CREATE PROCEDURE`, uncomment the preserved legacy block immediately above it, re-run the (idempotent) installer.

### 8.1 `spGetIntakeFolderName` — `:1463-1498`

Simplest rewrite: single-row read, `tblDocumentRootDirectory.IntakeDirectory` → `tblDropboxRootConfig.IntakeDirectory`:
```sql
select IntakeDirectory as DocumentFolder
from dbo.tblDropboxRootConfig with (nolock)
where ConfigID = 1;
```
(`:1494-1496`). Legacy body fully preserved, commented, at `:1474-1484`.

### 8.2 `spGetDocumentFolderName` — `:1501-1646`

Open-case folder resolver — the canonical shape every other builder SP (8.3, 8.4, 8.5) follows. Reads the naming template + root, tokenizes the template via `fnGetListOfWords`, and assembles a dynamic-SQL `SELECT` that concatenates literal words with column lookups from `vwfrmClientLedger`/`tblDropD`, then executes it:

```sql
select @DocumentRootNaming = DocumentRootNaming, @TeamRootPath = TeamRootPath
from dbo.tblDropboxRootConfig with (nolock) where ConfigID = 1;
...
select @sql = 'SELECT REPLACE(LTRIM(RTRIM(''' + @TeamRootPath + ''' +';
declare cursorT cursor read_only for
select Word from dbo.fnGetListOfWords(@DocumentRootNaming, ' ') order by Position;
...
select @sql = left(@sql, len(@sql) - 1) + ' + ''' + @DocumentFolder + ''')), ''\'', ''/'') AS DocumentFolder ';
select @sql = @sql + ' FROM (select c.CaseID, c.Orig_Atty, d.CodeVal as Case_Letter,
                            c.Last_Name, c.First_Name, c.FileNo
                            from vwfrmClientLedger c (nolock)
                            inner join tblDropD d (nolock) on c.Case_Letter = d.Code
                            where d.FieldName = ''Case_Letter'') as X
                            WHERE CaseID = ' + convert(varchar(10), @CaseID);
exec (@sql);
```
(`:1594-1641`, condensed)

The load-bearing diff from the legacy body (preserved at `:1512-1574`) is exactly two things: (1) source table `tblDocumentRootDirectory` → `tblDropboxRootConfig`, and (2) the output wrap changes from bare `LTRIM(RTRIM(...))` to `REPLACE(LTRIM(RTRIM(...)), '\', '/')` (`:1630`) — the point where backslash-templated output becomes a forward-slash Dropbox path. Everything else (the tokenizer cursor loop, the `vwfrmClientLedger`/`tblDropD` join for the `Case_Letter` lookup) is unchanged from legacy.

### 8.3 `spGetClosedDocumentFolderName` — `:1649-1759`

Same pattern as 8.2, reading `DocumentClosedNaming` instead of `DocumentRootNaming` (`:1707-1710`), same `REPLACE(...,'\','/')` wrap (`:1743`). **Legacy body is abridged**, not fully preserved — see Discrepancies §1 below.

### 8.4 `spGetClosedFileScanFolderName` — `:1762-1875`

Same pattern, reading `ClosedFileScanNaming`/`ClosedFileScanDirectory` (`:1823-1826`), and its derived view additionally exposes `c.Yr` (`:1862`) because the closed-file-scan naming template uses `[Yr]` (per the seed value at `:144`, `N'\ TB \ [Yr] \'`). Same wrap at `:1859`. **Legacy body is abridged** — see Discrepancies §1.

### 8.5 `spGetAllInvoicesFolderName` — `:1878-1990`

Same overall pattern but with one structural addition not present in 8.2/8.3/8.4, because `tblDropboxRootConfig.AllInvoicesNaming` is seeded empty (`:144`):

```sql
-- AllInvoicesNaming is empty in current config; the cursor still runs
-- (zero iterations) and the SQL falls through to the constant root path.
...
-- If AllInvoicesNaming was empty, @sql ends with '''+' (no trailing +)
-- and LEFT(...,LEN-1) would corrupt the SQL. Guard:
if right(@sql, 1) = '+'
    select @sql = left(@sql, len(@sql) - 1);

select @sql = @sql + ' + ''' + isnull(@DocumentFolder, '') + ''')), ''\'', ''/'') AS DocumentFolder ';
```
(`:1940-1974`)

Every other builder SP unconditionally does `left(@sql, len(@sql) - 1)` to drop a trailing `+` that the cursor loop is assumed to have appended — safe only because their naming templates are non-empty. `spGetAllInvoicesFolderName` is the one SP whose naming template is empty by design, so it needs (and gets) a conditional guard plus an `isnull()` wrap around `@DocumentFolder`. **Legacy body is abridged** — see Discrepancies §1.

### 8.6 `spMoveDocumentFolder` (G2 rewrite) — `:1993-2126`

Signature change from legacy `(@CaseID, @CaseStatus)` to `(@CaseID, @OldFolderPath, @NewFolderPath)` (`:2066-2069`) — the caller (VBA) now computes both paths itself (via 8.2/8.3) rather than the SP inferring a `_CLOSED` insertion point by token position. Normalizes both paths to a trailing slash (`:2076-2077`), then updates both tables in one transaction:

```sql
set xact_abort on;
...
update dbo.tblCaseDocuments
set DocumentFileName = @NewFolderPath + substring(DocumentFileName, len(@OldFolderPath) + 1, len(DocumentFileName))
where CaseID = @CaseID and left(DocumentFileName, len(@OldFolderPath)) = @OldFolderPath;
set @CaseDocsUpdated = @@rowcount;

update dbo.tblScans
set ScanLocation = @NewFolderPath + substring(ScanLocation, len(@OldFolderPath) + 1, len(ScanLocation))
where CaseID = @CaseID and left(ScanLocation, len(@OldFolderPath)) = @OldFolderPath;
set @ScansUpdated = @@rowcount;

if @CaseDocsUpdated = 0 and @ScansUpdated = 0
begin
    rollback tran;
    ;throw 51000, N'spMoveDocumentFolder: zero rows updated in both tblCaseDocuments and tblScans ...', 1;
end
commit tran;
```
(`:2073-2109`, condensed)

Safety guards: `set xact_abort on` (`:2073`) so any runtime error aborts the whole batch; explicit `begin try/catch` with `rollback tran` + re-`throw` (`:2081`, `:2111-2114`); the "both-zero-rowcount hard-fail" (`:2098-2107`) is the load-bearing guard — a case having a row updated in *only one* of the two tables is accepted (common; comment at `:2101-2102` explains many cases legitimately live in only one table), but both-zero means the SQL ledger has no record of the case under `@OldFolderPath` at all, which the SP treats as a sign of a path mismatch that must not be silently accepted.

Cross-reference (secondary source): `DocumentManagement.bas:1409-1412` strips the trailing slash from both paths before calling `DropboxService.MoveFile` (Dropbox `move_v2` rejects a trailing slash), then passes the **same** (slash-stripped) path variables into the SP call at `:1430-1433` — the SP re-adds the trailing slash itself at `:2076-2077` for its own prefix-match purposes. Two callers, two conventions, both intentional (`:2075` in the SQL file documents the SP's own normalization step). If the SP throws (its 51000 error), `DocumentManagement.bas:1442-1471` compensates by reverse-moving the Dropbox folder and either confirms the revert or surfaces a "CRITICAL … contact IT" message if the revert also fails.

### 8.7 `spSaveCaseDocument` (G13 rewrite) — `:2129-2201`

Adds an unresolved-template-token guard ahead of the legacy delete+insert body:

```sql
if @DocumentName like '%[[]%]%'
   or charindex('<currentdate>', @DocumentName) > 0
   or charindex('(customuserentry)', @DocumentName) > 0
begin
    ;throw 51001, N'spSaveCaseDocument: @DocumentName contains unresolved template tokens ...', 1;
end;
```
(`:2178-2183`)

This is the runtime enforcement counterpart to the historical typo fixed in Section 3 (`:387-390`) and to the 9-row "unresolved template" defect class surfaced by the Section 7.2 listing (`:1287-1293`) — it prevents new rows from being saved with a literal `[Case_Letter]`-style token, an unsubstituted `<currentdate>`, or the (now-corrected) `(customuserentry)` marker still embedded in the stored filename. Comment at `:2136-2138` notes it deliberately does *not* block ordinary parentheses/dashes in filenames — only these three specific unsubstituted-token shapes.

### 8.8 Verification — `:2204-2240`

Not a rewrite; runs each of the five builder SPs against the known-good QUAIL/case 30337 and `print`s a header before each so IT can eyeball the resolved forward-slash path:
```sql
exec dbo.spGetIntakeFolderName;
exec dbo.spGetDocumentFolderName 'General', 30337;
exec dbo.spGetClosedDocumentFolderName 'General', 30337;
exec dbo.spGetClosedFileScanFolderName 'General', 30337;
exec dbo.spGetAllInvoicesFolderName 30337;
```
(`:2216-2236`)

---

## Secondary/cross-reference sources (VBA)

These confirm what the SQL-side path objects are *for*, from the calling side. Not the primary source for this task — cited only to corroborate the SQL findings above.

- `DocumentManagement.bas:89`, `:121`, `:153`, `:183`, `:339`, `:371`, `:802`, `:1695` — the `exec sp...` call sites for `spGetDocumentFileName`, `spGetDocumentFolderName`, `spGetIntakeFolderName`, `spGetClosedDocumentFolderName`, `spGetClosedFileScanFolderName`, `spGetAllInvoicesFolderName`, `spSaveCaseDocument`, `spGetIntakeDocumentFileName` respectively — confirms each SP named in the SQL file's Section 8 has exactly one live VBA caller.
- `DocumentManagement.bas:261` — `GetDocumentRootFolder()` reads `tblDropboxRootConfig.TeamRootPath` directly (`SELECT TeamRootPath FROM dbo.tblDropboxRootConfig WHERE ConfigID = 1`) and feeds it to `DropboxService.DropboxPathToLocalPath` — confirms `TeamRootPath` (seeded in SQL Section 1.1) is consumed on the VBA side, independent of the builder SPs.
- `DocumentManagement.bas:1426-1433` — the live call site for the rewritten `spMoveDocumentFolder`, matching the new `(@CaseID, @OldFolderPath, @NewFolderPath)` signature from SQL `:2066-2069` exactly.
- `DocumentManagement.bas:45-46` — Phase 4 status header states "Phase 4c — Stored procedure rewrites (DONE in Dropbox-Migration-SQL-Install.sql Section 8). Live on awsql2022dev/TateByWater," directly confirming this SQL file's Section 8 is the live, applied version of these SPs in the test environment.
- `DropboxService.bas:4315-4334` (`LocalPathToDropboxPath`) and `:4343-4357` (`DropboxPathToLocalPath`) — VBA-side string-manipulation helpers that convert between a Windows local-synced path and a `/Company/...` Dropbox API path (simple prefix-swap + `\`/`/` `Replace()`, no SQL involved). These are a *different* path concern than anything in the SQL file: they translate between "Dropbox path" and "local filesystem mirror of that Dropbox path," not between "legacy `S:\` path" and "Dropbox path." Kept distinct in this doc so the two aren't conflated.
- `DropboxService.bas:607` — builds the `Dropbox-API-Path-Root` header value from `m_NamespaceId`, matching `tblDropboxRootConfig.NamespaceId` seeded at SQL `:141`.

---

## Discrepancies vs. CLAUDE.md's section map

1. **Section 9 exists and is undocumented — in both CLAUDE.md and the SQL file's own header.** `Dropbox-Migration-SQL-Install.sql:2243-2322` is `SECTION 9 — DROPBOX BRIDGE SERVICE OBJECTS`, adding a `BridgeUrl` column to `tblDropboxConfig` (`:2267-2287`) and a new `tblDropboxServiceToken` table (`:2292-2319`). CLAUDE.md's section map ("SQL installer section map") lists only Sections 1–8 and never mentions a 9th. This isn't just a CLAUDE.md gap: the SQL file's *own* top-of-file "WHAT THIS SCRIPT DOES" comment block (`:1-83`) also stops at Section 8 (`:31-41`) and never mentions Section 9 — so the file's internal map is out of date, not just CLAUDE.md's summary of it. Section 9 is not itself document-path migration (`BridgeUrl` is a service endpoint URL, not a case-document path), so its absence from a *path-migration-focused* map is defensible — but its absence from the file's own complete section inventory is a genuine drift worth flagging. One related detail CLAUDE.md's DESTRUCTIVE warning (which does list `tblDropboxConfig`'s `AppSecret` reset) doesn't mention: the re-run caveat at `:2254-2257` notes that because Section 1.2 drops+recreates `tblDropboxConfig`, a destructive re-run also drops the `BridgeUrl` column, which Section 9.1 then silently re-adds with a placeholder URL — a second "IT must re-apply after every run" item beyond just `AppSecret`.

2. **The documented single-file rollback procedure does not work for three of the seven rewritten SPs.** The Section 8 header states the rollback procedure as: "comment out the new create-procedure, uncomment the legacy create-procedure, re-run this installer" (`:1446-1449`). This is accurate for `spGetIntakeFolderName` (8.1, full legacy body at `:1474-1484`), `spGetDocumentFolderName` (8.2, full legacy body at `:1512-1574`), `spMoveDocumentFolder` (8.6, full legacy body at `:2007-2061`), and `spSaveCaseDocument` (8.7, full legacy body at `:2144-2160`). It is **not** accurate for `spGetClosedDocumentFolderName` (8.3), `spGetClosedFileScanFolderName` (8.4), and `spGetAllInvoicesFolderName` (8.5): their "LEGACY" blocks are abridged, each containing a placeholder comment in place of the actual tokenizer-cursor logic —
   - `:1682`: `-- ... [tokenizer cursor; same shape as 8.2 LEGACY] ...`
   - `:1795`: `-- ... [tokenizer cursor; same shape as 8.2 LEGACY] ...`
   - `:1909`: `-- ... [tokenizer cursor; same shape as 8.2 LEGACY] ...`

   Uncommenting any of these three as instructed would produce a `CREATE PROCEDURE` body with a syntactically incomplete `@sql`-building block (the `DECLARE cursorT ... OPEN ... WHILE ... END` loop is missing entirely, only a summary comment stands in for it), which would not compile. Rolling back 8.3, 8.4, or 8.5 by the documented procedure would require reconstructing the omitted cursor logic from the full pattern shown in 8.2's complete legacy block (`:1512-1574`) — not a same-file, no-thought operation as the header implies for the other four SPs. This is a real gap between the stated rollback safety guarantee and what the file actually contains for these three procedures.

3. **Minor: the "11 per-category passes" figure is correct but only resolves cleanly once the two quarantine inserts are excluded from the count.** CLAUDE.md's "11 per-category passes" description is accurate (Part A: 1, Part B rewrites B-1..B-6: 6, Part C rewrites C-1..C-4: 4 → 11), and matches the file's own section comments (`:500` "single... pass" for Part A, `:538` "six per-category passes + quarantine" for Part B, `:742` "four passes + quarantine" for Part C) — but a reader who counts every labeled sub-pass, including the quarantine `INSERT`s (B-7 at `:714-739`, C-5 at `:828-856`), would get 13, not 11. Worth stating explicitly since the file itself doesn't total the count anywhere. Also note the quarantine steps are labeled inconsistently — Part B's is "B-7" as if a 7th rewrite pass, Part C's is "C-5" — though this is a labeling nit, not a functional issue.
