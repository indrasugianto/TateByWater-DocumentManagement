-- =============================================================================
-- Dropbox-Migration-PROD-UNDO.sql
--
-- PURPOSE
--   Reverse, on the PRODUCTION database [TateBywater] (host tbf-cms), every
--   change made by an ACCIDENTAL run of Dropbox-Migration-SQL-Install.sql.
--   The accident's object timestamps are 2026-06-02 ~16:23.
--
-- SOURCE OF TRUTH
--   [TateByWater_Restore] — a restore of the 2026-06-01 12:09 PM backup,
--   taken BEFORE the accident. This script reads original object definitions
--   and original row values from it.
--
-- WHY THIS IS SURGICAL (NOT A WHOLESALE TABLE RESTORE)
--   Legitimate production activity occurred between the 06-01 backup and the
--   06-02 accident:
--     - tblCaseDocuments : 32 rows inserted today (PKs 32386-32417),
--                          5 backup rows deleted today (normal app activity).
--     - [TB Intakes]     : 14 rows inserted today (PKs 2693-2706).
--     - tblScans         : no inserts/deletes.
--   A wholesale restore would destroy those new rows and resurrect deleted
--   ones. Instead we restore ONLY the columns the installer rewrote, matched
--   by primary key, and we reverse the path rewrite on the new-today rows
--   algorithmically (they have no counterpart in the backup).
--
-- WHAT THE INSTALLER CHANGED (and what this script does to each):
--   1. CREATED 8 tables + spLogDropboxAuditEvent  -> DROP them (guarded:
--      only dropped if absent from the restore DB, proving they are new).
--   2. OVERWROTE 7 stored procedures              -> recreate from restore.
--   3. tblDocumentTypes row 30 DocumentNamingRule -> restore from backup.
--   4. [TB Intakes] GI Last/First for IDs 183,185,193,2482 -> restore.
--   5. tblCaseDocuments.DocumentFileName          -> restore matched rows;
--                                                    reverse-transform new.
--      tblScans.ScanLocation                      -> restore matched rows.
--      [TB Intakes].[Scan Location GI]            -> restore matched rows;
--                                                    reverse-transform new.
--
-- WHAT THIS SCRIPT DELIBERATELY DOES NOT TOUCH
--   - The 5 case-document rows deleted today (not caused by the installer).
--   - Any row inserted today other than to reverse the installer's path
--     rewrite on it.
--   - fnGetListOfWords, spGetCaseDocument, spGetDocumentFileName,
--     spGetIntakeDocumentFileName, tblDocumentRootDirectory (installer left
--     these alone).
--
-- SAFETY
--   - set xact_abort on; the data restore runs in ONE transaction and only
--     commits if the post-restore leftover-offender count = 0.
--   - Idempotent: re-running is a no-op (difference-guarded UPDATEs, DROP IF
--     EXISTS, LIKE-guarded reversals, restore-presence guards on drops).
--   - Run as a single batch set in SSMS/sqlcmd. Review SECTION 0 output and
--     SECTION 5 verification before trusting the result.
--
-- PRE-FLIGHT
--   Confirm both databases are on the same instance and online:
--     SELECT name, state_desc FROM sys.databases
--      WHERE name IN ('TateBywater','TateByWater_Restore');
-- =============================================================================

set nocount on;
set xact_abort on;
go

use TateBywater;
go

declare @prod sysname = db_name();
print N'================================================================================';
print N'PROD-UNDO starting against database: ' + @prod;
print N'Restore source: TateByWater_Restore';
print N'================================================================================';
go

-- Hard stop if we are not pointed at the intended production DB, or the
-- restore DB is missing. Prevents running this against the wrong target.
if db_name() <> N'TateBywater'
    throw 60000, N'Refusing to run: current database is not TateBywater.', 1;
if db_id(N'TateByWater_Restore') is null
    throw 60001, N'Refusing to run: TateByWater_Restore is not present on this instance.', 1;
go


-- #############################################################################
-- SECTION 0 — PRE-CHANGE SNAPSHOT (output only)
-- #############################################################################
print N'';
print N'>>> SECTION 0: pre-change snapshot';
go

select 'objects-created-by-installer (expect 9, all present=damage)' as check_name,
       count(*) as n
from sys.objects
where name in (N'tblDropboxRootConfig', N'tblDropboxConfig', N'tblDropboxRevocationList',
               N'tblDropboxAuditLog', N'tblDropboxOrphanQueue', N'tblDropboxVerificationReport',
               N'tblScans_ManualTriage', N'TBIntakes_ManualTriage', N'spLogDropboxAuditEvent');

select 'tblCaseDocuments paths starting /Company/ (damage)' as check_name, count(*) as n
from dbo.tblCaseDocuments where DocumentFileName like '/Company/%'
union all
select 'tblScans paths starting /Company/ (damage)', count(*)
from dbo.tblScans where ScanLocation like '/Company/%'
union all
select 'TB Intakes paths starting /Company/ (damage)', count(*)
from dbo.[TB Intakes] where [Scan Location GI] like '/Company/%';
go


-- #############################################################################
-- SECTION 1 — DROP INSTALLER-CREATED OBJECTS
--   Guarded: each object is dropped ONLY if it does NOT exist in the restore
--   DB. Existence in the restore (pre-accident) backup would prove the object
--   pre-dated the accident and must be kept — in that case we skip the drop
--   and print a warning.
-- #############################################################################
print N'';
print N'>>> SECTION 1: drop installer-created objects';
go

declare @newobjs table (obj sysname, kind char(1)); -- kind: T=table, P=proc
insert into @newobjs (obj, kind) values
    (N'tblDropboxVerificationReport', 'T'),
    (N'tblDropboxOrphanQueue',        'T'),
    (N'tblDropboxAuditLog',           'T'),
    (N'tblDropboxRevocationList',     'T'),
    (N'tblDropboxConfig',             'T'),
    (N'tblDropboxRootConfig',         'T'),
    (N'tblScans_ManualTriage',        'T'),
    (N'TBIntakes_ManualTriage',       'T'),
    (N'spLogDropboxAuditEvent',       'P');

declare @obj sysname, @kind char(1), @sql nvarchar(max), @existsInRestore int;
declare obj_cur cursor local fast_forward for select obj, kind from @newobjs;
open obj_cur;
while 1 = 1
begin
    fetch obj_cur into @obj, @kind;
    if @@fetch_status <> 0 break;

    set @existsInRestore =
        case when object_id(N'TateByWater_Restore.dbo.' + quotename(@obj)) is not null then 1 else 0 end;

    if @existsInRestore = 1
    begin
        print N'    SKIP (exists in restore — pre-dates accident, keeping): ' + @obj;
        continue;
    end;

    if object_id(N'dbo.' + quotename(@obj)) is null
    begin
        print N'    already absent: ' + @obj;
        continue;
    end;

    set @sql = case @kind
                   when 'T' then N'drop table dbo.' + quotename(@obj) + N';'
                   when 'P' then N'drop procedure dbo.' + quotename(@obj) + N';'
               end;
    exec sys.sp_executesql @sql;
    print N'    dropped ' + case @kind when 'T' then N'table ' else N'procedure ' end + @obj;
end;
close obj_cur;
deallocate obj_cur;
go
print N'    SECTION 1 done.';
go


-- #############################################################################
-- SECTION 2 — RESTORE ORIGINAL STORED PROCEDURE DEFINITIONS
--   Pulls the pre-accident CREATE PROCEDURE text from TateByWater_Restore and
--   recreates each proc verbatim in production. No per-proc grants existed in
--   the backup, so none are re-applied.
-- #############################################################################
print N'';
print N'>>> SECTION 2: restore original stored procedures from backup';
go

declare @procs table (nm sysname);
insert into @procs (nm) values
    (N'spGetIntakeFolderName'),
    (N'spGetDocumentFolderName'),
    (N'spGetClosedDocumentFolderName'),
    (N'spGetClosedFileScanFolderName'),
    (N'spGetAllInvoicesFolderName'),
    (N'spMoveDocumentFolder'),
    (N'spSaveCaseDocument');

declare @nm sysname, @def nvarchar(max), @drop nvarchar(max);
declare p_cur cursor local fast_forward for select nm from @procs;
open p_cur;
while 1 = 1
begin
    fetch p_cur into @nm;
    if @@fetch_status <> 0 break;

    select @def = m.definition
    from TateByWater_Restore.sys.sql_modules m
    where m.object_id = object_id(N'TateByWater_Restore.dbo.' + quotename(@nm));

    if @def is null
    begin
        print N'    *** WARNING: no backup definition for ' + @nm + N' — left unchanged.';
        continue;
    end;

    set @drop = N'drop procedure if exists dbo.' + quotename(@nm) + N';';
    exec sys.sp_executesql @drop;       -- drop installer version
    exec sys.sp_executesql @def;        -- recreate original (its own batch)
    print N'    restored procedure ' + @nm;
end;
close p_cur;
deallocate p_cur;
go
print N'    SECTION 2 done.';
go


-- #############################################################################
-- SECTION 3 — RESTORE DATA (single transaction; commit only if clean)
-- #############################################################################
print N'';
print N'>>> SECTION 3: restore data';
go

begin try
    begin transaction UndoData;

    -- 3.1 tblDocumentTypes.DocumentNamingRule (installer SECTION 3 typo "fix")
    update p
    set p.DocumentNamingRule = r.DocumentNamingRule
    from dbo.tblDocumentTypes p
    join TateByWater_Restore.dbo.tblDocumentTypes r on r.DocumentTypeID = p.DocumentTypeID
    where exists (select p.DocumentNamingRule except select r.DocumentNamingRule);
    declare @n_doctypes int = @@rowcount;

    -- 3.2 [TB Intakes] GI Last/First Name (installer SECTION 4, IDs 183/185/193/2482)
    update p
    set p.[GI Last Name]  = r.[GI Last Name],
        p.[GI First Name] = r.[GI First Name]
    from dbo.[TB Intakes] p
    join TateByWater_Restore.dbo.[TB Intakes] r on r.[ID] = p.[ID]
    where exists (select p.[GI Last Name], p.[GI First Name]
                  except
                  select r.[GI Last Name], r.[GI First Name]);
    declare @n_giname int = @@rowcount;

    -- 3.3 tblCaseDocuments.DocumentFileName — matched rows restored from backup
    update p
    set p.DocumentFileName = r.DocumentFileName
    from dbo.tblCaseDocuments p
    join TateByWater_Restore.dbo.tblCaseDocuments r on r.CaseDocumentID = p.CaseDocumentID
    where exists (select p.DocumentFileName except select r.DocumentFileName);
    declare @n_docs_matched int = @@rowcount;

    -- 3.3b tblCaseDocuments — new-today rows (no backup counterpart): reverse
    --      the installer transform  S:\X\Y -> /Company/X/Y  back to  S:\X\Y.
    --      Guarded to rows the installer actually rewrote (LIKE '/Company/%').
    update dbo.tblCaseDocuments
    set DocumentFileName = 'S:\' + replace(substring(DocumentFileName, 10, len(DocumentFileName)), '/', '\')
    where CaseDocumentID not in
          (select r.CaseDocumentID from TateByWater_Restore.dbo.tblCaseDocuments r)
      and DocumentFileName like '/Company/%';
    declare @n_docs_new int = @@rowcount;

    -- 3.4 tblScans.ScanLocation — all rows matched by PK, restore from backup
    update p
    set p.ScanLocation = r.ScanLocation
    from dbo.tblScans p
    join TateByWater_Restore.dbo.tblScans r on r.ScansID = p.ScansID
    where exists (select p.ScanLocation except select r.ScanLocation);
    declare @n_scans int = @@rowcount;

    -- 3.5 [TB Intakes].[Scan Location GI] — matched rows restored from backup
    update p
    set p.[Scan Location GI] = r.[Scan Location GI]
    from dbo.[TB Intakes] p
    join TateByWater_Restore.dbo.[TB Intakes] r on r.[ID] = p.[ID]
    where exists (select p.[Scan Location GI] except select r.[Scan Location GI]);
    declare @n_intk_matched int = @@rowcount;

    -- 3.5b [TB Intakes] — new-today rows: reverse the installer transform.
    update dbo.[TB Intakes]
    set [Scan Location GI] = 'S:\' + replace(substring([Scan Location GI], 10, len([Scan Location GI])), '/', '\')
    where [ID] not in (select r.[ID] from TateByWater_Restore.dbo.[TB Intakes] r)
      and [Scan Location GI] like '/Company/%';
    declare @n_intk_new int = @@rowcount;

    -- -----------------------------------------------------------------------
    -- Leftover-offender check: after the undo, NO path column anywhere should
    -- still carry a /Company/ root (the installer's signature). New-today rows
    -- not matching '/Company/%' (e.g. NULL/blank/other) are reported in
    -- SECTION 4 for manual review but do not block commit.
    -- -----------------------------------------------------------------------
    declare @left_docs int, @left_scans int, @left_intk int;

    select @left_docs = count(*) from dbo.tblCaseDocuments where DocumentFileName like '/Company/%';
    select @left_scans = count(*) from dbo.tblScans where ScanLocation like '/Company/%';
    select @left_intk = count(*) from dbo.[TB Intakes] where [Scan Location GI] like '/Company/%';

    if @left_docs + @left_scans + @left_intk > 0
    begin
        rollback transaction UndoData;
        declare @m varchar(400) =
            N'PROD-UNDO aborted — /Company/ paths remain. tblCaseDocuments='
            + cast(@left_docs as varchar(10)) + N' tblScans=' + cast(@left_scans as varchar(10))
            + N' [TB Intakes]=' + cast(@left_intk as varchar(10))
            + N'. No changes committed.';
        throw 60010, @m, 1;
    end;

    commit transaction UndoData;

    print N'    SECTION 3 done. Rows changed:';
    print N'      tblDocumentTypes.DocumentNamingRule : ' + cast(@n_doctypes as varchar(10));
    print N'      [TB Intakes] GI Last/First Name     : ' + cast(@n_giname as varchar(10));
    print N'      tblCaseDocuments matched (from bkup) : ' + cast(@n_docs_matched as varchar(10));
    print N'      tblCaseDocuments new-today reversed  : ' + cast(@n_docs_new as varchar(10));
    print N'      tblScans matched (from backup)       : ' + cast(@n_scans as varchar(10));
    print N'      [TB Intakes] matched (from backup)   : ' + cast(@n_intk_matched as varchar(10));
    print N'      [TB Intakes] new-today reversed      : ' + cast(@n_intk_new as varchar(10));
end try
begin catch
    if xact_state() <> 0 rollback transaction UndoData;
    throw;
end catch;
go


-- #############################################################################
-- SECTION 4 — VERIFICATION (output only — expect zeros / clean)
-- #############################################################################
print N'';
print N'>>> SECTION 4: verification';
go

-- 4.1 Installer-created objects should all be gone.
select 'installer objects remaining (expect 0)' as check_name, count(*) as n
from sys.objects
where name in (N'tblDropboxRootConfig', N'tblDropboxConfig', N'tblDropboxRevocationList',
               N'tblDropboxAuditLog', N'tblDropboxOrphanQueue', N'tblDropboxVerificationReport',
               N'tblScans_ManualTriage', N'TBIntakes_ManualTriage', N'spLogDropboxAuditEvent');

-- 4.2 No /Company/ paths anywhere (expect 0 each).
select 'tblCaseDocuments /Company/ remaining' as check_name, count(*) as n
from dbo.tblCaseDocuments where DocumentFileName like '/Company/%'
union all
select 'tblScans /Company/ remaining', count(*)
from dbo.tblScans where ScanLocation like '/Company/%'
union all
select 'TB Intakes /Company/ remaining', count(*)
from dbo.[TB Intakes] where [Scan Location GI] like '/Company/%';

-- 4.3 The 7 procs should match the backup definitions byte-for-byte.
select 'procs differing from backup (expect 0)' as check_name, count(*) as n
from (values (N'spGetIntakeFolderName'),(N'spGetDocumentFolderName'),
             (N'spGetClosedDocumentFolderName'),(N'spGetClosedFileScanFolderName'),
             (N'spGetAllInvoicesFolderName'),(N'spMoveDocumentFolder'),
             (N'spSaveCaseDocument')) v(nm)
where isnull((select m.definition from sys.sql_modules m where m.object_id = object_id(N'dbo.' + quotename(v.nm))), '~A~')
   <> isnull((select m.definition from TateByWater_Restore.sys.sql_modules m
              where m.object_id = object_id(N'TateByWater_Restore.dbo.' + quotename(v.nm))), '~B~');

-- 4.4 Section 4 intake rows should match backup again.
select p.[ID], p.[GI Last Name] as last_now, p.[GI First Name] as first_now
from dbo.[TB Intakes] p where p.[ID] in (183,185,193,2482) order by p.[ID];

-- 4.5 tblDocumentTypes row 30 should read the original (typo) rule again.
select DocumentTypeID, DocumentNamingRule from dbo.tblDocumentTypes where DocumentTypeID = 30;

-- 4.6 New-today rows that could NOT be auto-reversed (path was not /Company/-
--     rooted). Expect 0; any rows here need manual review.
select 'tblCaseDocuments new-today not auto-reversed' as check_name, count(*) as n
from dbo.tblCaseDocuments p
where p.CaseDocumentID not in (select r.CaseDocumentID from TateByWater_Restore.dbo.tblCaseDocuments r)
  and (p.DocumentFileName is null or (p.DocumentFileName not like 'S:\%' and p.DocumentFileName <> ''))
union all
select 'TB Intakes new-today not auto-reversed', count(*)
from dbo.[TB Intakes] p
where p.[ID] not in (select r.[ID] from TateByWater_Restore.dbo.[TB Intakes] r)
  and (p.[Scan Location GI] is null or (p.[Scan Location GI] not like 'S:\%' and p.[Scan Location GI] <> ''));
go

print N'';
print N'================================================================================';
print N'PROD-UNDO complete. Review SECTION 4 output — all counts should be 0.';
print N'================================================================================';
go
