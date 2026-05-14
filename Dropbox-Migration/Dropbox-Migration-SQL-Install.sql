-- =============================================================================
-- INSTALL_all.sql — Consolidated Phase 1b + Phase 2 installer for the
-- TBCMS Dropbox migration. Single end-to-end script.
--
-- See .docs/dropbox-migration-plan.md for the full plan.
--
-- WHAT THIS SCRIPT DOES (in order):
--   SECTION 1 — Phase 2 schema      : DROP+CREATE the 6 Dropbox tables and
--                                     spLogDropboxAuditEvent; seed singleton
--                                     config rows.
--   SECTION 2 — Manual-triage tables: CREATE-IF-MISSING the two triage tables
--                                     used by the path migration.
--   SECTION 3 — DocumentTypes typo  : fix the '(customeruserentry)' typo on
--                                     tblDocumentTypes row 30.
--   SECTION 4 — Intake natural-key  : fix [TB Intakes] rows 183/185/193/2482
--                                     (Phase 1b item B-5 — see plan).
--   SECTION 5 — Path migration      : rewrite S:\ paths to /Company/ across
--                                     tblCaseDocuments, tblScans, [TB Intakes]
--                                     in 11 per-category passes; quarantine
--                                     unrewritable rows; auto-commit only if
--                                     leftover-offender count = 0.
--   SECTION 6 — Verification        : leftover-offender check, prefix
--                                     distribution, sample of migrated rows.
--   SECTION 7 — Phase 1b listings   : diagnostic output for items the IT
--                                     admin must triage by hand —
--                                       B-9 path-length pre-flight
--                                       B-4 13 bracket-literal rows
--                                       B-6 198 non-canonical roots
--                                       B-7 tblScans manual-triage queue
--                                       B-8 [TB Intakes] manual-triage queue
--
-- TARGETS  : awsql2022dev/TateByWater (test env, Phases 0–6) and the
--            production SQL host at Phase 7 cutover.
--            Environment-agnostic — only USE TateByWater binds it.
--
-- IDEMPOTENT: yes. Every section is safe to re-run:
--   - Schema uses DROP IF EXISTS + CREATE.
--   - Data-quality fixes have WHERE guards so re-runs are no-ops.
--   - Path migration filters on legacy patterns ('S:\', '#…#') that won't
--     match post-rewrite paths.
--   - Manual-triage tables use CREATE-IF-MISSING + dedupe inserts so prior
--     triage rows survive re-runs.
--
-- *** DESTRUCTIVE — READ BEFORE RE-RUNNING ON A NON-FRESH DATABASE ***
--   Re-running this script DROPS the 6 Dropbox tables before recreating
--   them, which wipes:
--     - tblDropboxAuditLog          : ALL audit history.
--     - tblDropboxOrphanQueue       : queued orphan rows.
--     - tblDropboxVerificationReport: report history.
--     - tblDropboxRevocationList    : revocation entries.
--     - tblDropboxConfig            : AppSecret reset to placeholder
--                                     (IT must re-UPDATE after every run).
--     - tblDropboxRootConfig        : seed values reset to plan-defaults.
--   The manual-triage tables and the source tables (tblCaseDocuments,
--   tblScans, [TB Intakes], tblDocumentTypes) are NOT dropped.
--
-- POST-INSTALL ACTIONS (always required):
--   1. UPDATE dbo.tblDropboxConfig SET AppSecret = N'<real secret from
--      Dropbox App Console>' WHERE ConfigID = 1;
--   2. Review SECTION 7 listing output and author per-row fixes for items
--      B-4, B-6, B-7, B-8 before Phase 7 verification runs.
-- =============================================================================

use TateByWater;
go

set xact_abort on;
set nocount on;
go

print N'================================================================================';
print N'INSTALL_all.sql — starting against database: ' + db_name();
print N'================================================================================';
go


-- #############################################################################
-- SECTION 1 — PHASE 2 SCHEMA (drop + create)
-- #############################################################################

print N'';
print N'>>> SECTION 1: Phase 2 schema — drop + create';
go

-- Reverse dependency order: SP first, then tables.
drop procedure if exists dbo.spLogDropboxAuditEvent;
go

drop table if exists dbo.tblDropboxVerificationReport;
drop table if exists dbo.tblDropboxOrphanQueue;
drop table if exists dbo.tblDropboxAuditLog;
drop table if exists dbo.tblDropboxRevocationList;
drop table if exists dbo.tblDropboxConfig;
drop table if exists dbo.tblDropboxRootConfig;
go

-- ---------------------------------------------------------------------------
-- 1.1 tblDropboxRootConfig — per-case-folder template config.
-- ---------------------------------------------------------------------------
create table dbo.tblDropboxRootConfig
(
    ConfigID int not null primary key,
    NamespaceId varchar(50) not null,
    TeamRootPath varchar(500) not null,
    DocumentRootNaming varchar(500) not null,
    DocumentClosedNaming varchar(500) not null,
    AllInvoicesDirectory varchar(500) not null,
    AllInvoicesNaming varchar(500) not null,
    ClosedFileScanDirectory varchar(500) not null,
    ClosedFileScanNaming varchar(500) not null,
    ScannerDirectory varchar(500) not null,
    IntakeDirectory varchar(500) not null,
    constraint CK_DropboxRootConfig_SingleRow check (ConfigID = 1)
);
go

insert into dbo.tblDropboxRootConfig
(
    ConfigID,
    NamespaceId,
    TeamRootPath,
    DocumentRootNaming,
    DocumentClosedNaming,
    AllInvoicesDirectory,
    AllInvoicesNaming,
    ClosedFileScanDirectory,
    ClosedFileScanNaming,
    ScannerDirectory,
    IntakeDirectory
)
values
(1, N'14334595683', N'/Company/COMMON',
 N'\ [Orig_Atty] \_CLIENTS\ [Case_Letter] \ [Last_Name] , ~ [First_Name] ~ [FileNo] \',
 N'\ [Orig_Atty] \_CLIENTS\ [Case_Letter] \ _CLOSED \ [Last_Name] , ~ [First_Name] ~ [FileNo] \',
 N'/Company/COMMON/_ALL INVOICES', N'', N'/Company/Closed File Scans', N'\ TB \ [Yr] \', N'/Company/COMMON/_SCANNER',
 N'/Company/COMMON/Intakes');
go

-- ---------------------------------------------------------------------------
-- 1.2 tblDropboxConfig — App credentials + OAuth redirect URI (singleton).
--     AppSecret is plaintext; protection boundary is SQL credential.
--     IT MUST UPDATE the AppSecret post-install — placeholder is intentional.
-- ---------------------------------------------------------------------------
create table dbo.tblDropboxConfig
(
    ConfigID int not null primary key,
    AppKey varchar(200) not null,
    AppSecret varchar(200) not null,
    RedirectUri varchar(500) not null,
    constraint CK_DropboxConfig_SingleRow check (ConfigID = 1)
);
go

insert into dbo.tblDropboxConfig
(
    ConfigID,
    AppKey,
    AppSecret,
    RedirectUri
)
values
(1, N'dqleswbnux8k3m5', N'REPLACE_WITH_APP_SECRET_FROM_DROPBOX_APP_CONSOLE', N'http://localhost:8765');
go

-- ---------------------------------------------------------------------------
-- 1.3 tblDropboxRevocationList — IT-admin-managed list of revoked accounts.
-- ---------------------------------------------------------------------------
create table dbo.tblDropboxRevocationList
(
    RevocationID int not null identity(1, 1) primary key,
    DropboxAccountEmail varchar(320) not null,
    RevokedAt datetime not null,
    RevokedBy varchar(200) null,
    Reason varchar(500) null
);
go

create index IX_DropboxRevocationList_Email
on dbo.tblDropboxRevocationList (DropboxAccountEmail);
go

-- ---------------------------------------------------------------------------
-- 1.4 tblDropboxAuditLog — audit-critical Dropbox events.
-- ---------------------------------------------------------------------------
create table dbo.tblDropboxAuditLog
(
    AuditID int not null identity(1, 1) primary key,
    EventDate datetime not null
        constraint DF_DropboxAuditLog_EventDate
            default (sysdatetime()),
    DropboxAccountEmail varchar(320) null,
    CaseID int null,
    DocumentType varchar(100) null,
    DropboxPath varchar(max) null,
    ActionType varchar(50) not null,
    Outcome varchar(20) not null,
    ErrorDetail varchar(max) null,
    constraint CK_DropboxAuditLog_ActionType check (ActionType in ( N'Upload', N'Move', N'Copy', N'Delete',
                                                                    N'LinkGenerate'
                                                                  )
                                                   ),
    constraint CK_DropboxAuditLog_Outcome check (Outcome in ( N'Success', N'Failure' ))
);
go

create index IX_DropboxAuditLog_EventDate
on dbo.tblDropboxAuditLog (EventDate);
go

create index IX_DropboxAuditLog_AccountEmail_EventDate
on dbo.tblDropboxAuditLog (
                              DropboxAccountEmail,
                              EventDate
                          );
go

-- ---------------------------------------------------------------------------
-- 1.5 tblDropboxOrphanQueue — failed upload-compensation queue.
-- ---------------------------------------------------------------------------
create table dbo.tblDropboxOrphanQueue
(
    OrphanID int not null identity(1, 1) primary key,
    EventDate datetime not null
        constraint DF_DropboxOrphanQueue_EventDate
            default (sysdatetime()),
    DropboxAccountEmail varchar(320) null,
    OrphanDropboxPath varchar(max) not null,
    WorkflowName varchar(50) not null,
    CaseID int null,
    DocumentType varchar(100) null,
    OriginalSPError varchar(max) not null,
    CompensatingDeleteError varchar(max) not null,
    Resolution varchar(20) not null
        constraint DF_DropboxOrphanQueue_Resolution
            default (N'Open'),
    ResolvedAt datetime null,
    ResolvedBy varchar(200) null,
    ResolutionNote varchar(max) null,
    constraint CK_DropboxOrphanQueue_Resolution check (Resolution in ( N'Open', N'RetriedSP', N'DeletedManually',
                                                                       N'KeptAsIs'
                                                                     )
                                                      )
);
go

create index IX_DropboxOrphanQueue_Resolution
on dbo.tblDropboxOrphanQueue (Resolution)
include (
            EventDate,
            WorkflowName
        );
go

-- ---------------------------------------------------------------------------
-- 1.6 tblDropboxVerificationReport — pre-cutover gate artifact.
--     SourceRowID is the surrogate PK for all 3 source tables:
--       tblCaseDocuments → CaseDocumentID
--       tblScans         → ScansID
--       TB Intakes       → [ID] (confirmed PK 2026-05-14)
-- ---------------------------------------------------------------------------
create table dbo.tblDropboxVerificationReport
(
    VerificationID int not null identity(1, 1) primary key,
    SourceTable varchar(50) not null,
    SourceRowID int not null,
    DropboxPath varchar(max) not null,
    Status varchar(20) not null,
    ErrorDetail varchar(max) null,
    CheckedAt datetime not null
        constraint DF_DropboxVerificationReport_CheckedAt
            default (sysdatetime()),
    constraint CK_VerificationReport_SourceTable check (SourceTable in ( N'tblCaseDocuments', N'tblScans',
                                                                         N'TB Intakes'
                                                                       )
                                                       ),
    constraint CK_VerificationReport_Status check (Status in ( N'Found', N'NotFound', N'Error' ))
);
go

create index IX_DropboxVerificationReport_Status_SourceTable
on dbo.tblDropboxVerificationReport (
                                        Status,
                                        SourceTable
                                    );
go

-- ---------------------------------------------------------------------------
-- 1.7 spLogDropboxAuditEvent — single insert helper for tblDropboxAuditLog.
-- ---------------------------------------------------------------------------
go
create procedure dbo.spLogDropboxAuditEvent
    @DropboxAccountEmail varchar(320) = null,
    @CaseID int = null,
    @DocumentType varchar(100) = null,
    @DropboxPath varchar(max) = null,
    @ActionType varchar(50),
    @Outcome varchar(20),
    @ErrorDetail varchar(max) = null
as
begin
    set nocount on;
    insert into dbo.tblDropboxAuditLog
    (
        EventDate,
        DropboxAccountEmail,
        CaseID,
        DocumentType,
        DropboxPath,
        ActionType,
        Outcome,
        ErrorDetail
    )
    values
    (sysdatetime(), @DropboxAccountEmail, @CaseID, @DocumentType, @DropboxPath, @ActionType, @Outcome, @ErrorDetail);
end;
go

print N'    SECTION 1 done.';
go


-- #############################################################################
-- SECTION 2 — MANUAL-TRIAGE TABLES (create if missing — preserved across runs)
-- #############################################################################

print N'';
print N'>>> SECTION 2: manual-triage tables';
go

if object_id('dbo.tblScans_ManualTriage', 'U') is null
begin
    create table dbo.tblScans_ManualTriage
    (
        ScansID int not null primary key,
        OriginalValue varchar(max) not null,
        Reason varchar(200) not null,
        QuarantinedAt datetime not null
            default sysdatetime()
    );
end;

if object_id('dbo.TBIntakes_ManualTriage', 'U') is null
begin
    -- [TB Intakes] does have a surrogate PK ([ID]); the natural-key columns
    -- here are kept for triage convenience, not because they're required.
    create table dbo.TBIntakes_ManualTriage
    (
        TriageID int identity(1, 1) primary key,
        IntakeID int null,
        GILastName varchar(255) null,
        GIFirstName varchar(255) null,
        GIDate datetime null,
        OriginalValue varchar(max) not null,
        Reason varchar(200) not null,
        QuarantinedAt datetime not null
            default sysdatetime()
    );
end;
go

print N'    SECTION 2 done.';
go


-- #############################################################################
-- SECTION 3 — PHASE 1b: DocumentTypes typo fix
--   tblDocumentTypes row 30 ('General') has '(customeruserentry)' (extra 'e').
--   The path tokenizer recognises '(customuserentry)'. Fix is idempotent.
-- #############################################################################

print N'';
print N'>>> SECTION 3: Phase 1b — DocumentTypes naming-rule typo';
go

begin try
    begin transaction;

    update dbo.tblDocumentTypes
    set DocumentNamingRule = replace(DocumentNamingRule, '(customeruserentry)', '(customuserentry)')
    where DocumentTypeID = 30
          and DocumentNamingRule like '%(customeruserentry)%';

    declare @DocTypeRowsUpdated int = @@rowcount;

    -- Sanity: confirm no row anywhere still carries the typo.
    declare @TypoStillPresent int;
    select @TypoStillPresent = count(*)
    from dbo.tblDocumentTypes
    where DocumentNamingRule like '%(customeruserentry)%';

    if @TypoStillPresent > 0
    begin
        rollback;
        throw 51001, N'SECTION 3 aborted: tblDocumentTypes still has rows containing the (customeruserentry) typo after fix attempt.', 1;
    end;

    commit;
    print N'    SECTION 3 done. RowsUpdated=' + cast(@DocTypeRowsUpdated as varchar(10));
end try
begin catch
    if xact_state() <> 0
        rollback;
    throw;
end catch;
go


-- #############################################################################
-- SECTION 4 — PHASE 1b item B-5: intake natural-key fixes (4 confident rows)
--   The verification report uses [TB Intakes].[ID] (surrogate PK) so these
--   NULLs no longer block verification, but they still break Access form-
--   side lookups that filter by ([GI Last Name], [GI First Name], [GI Date]).
--
--   Decisions captured 2026-05-14:
--     ID 183 (JJM & Associates of VA) → business entity, [GI First Name] = ''
--     ID 185 → split crammed name to Last='Morales Bolanos', First='Marvin Eduardo'
--     ID 193 → fill Last='Carpio', First='Alejandro' from path
--     ID 2482 → fill Last='Kerr' from path
--   5 remaining rows have only a NULL [GI Date] with no recoverable signal;
--   they are left alone (non-blocker under the new verification design).
-- #############################################################################

print N'';
print N'>>> SECTION 4: Phase 1b B-5 — intake natural-key fixes';
go

begin try
    begin transaction;

    update dbo.[TB Intakes]
    set [GI First Name] = N''
    where [ID] = 183
          and [GI First Name] is null;

    update dbo.[TB Intakes]
    set [GI Last Name] = N'Morales Bolanos',
        [GI First Name] = N'Marvin Eduardo'
    where [ID] = 185
          and
          (
              [GI Last Name] = N'Morales Bolanos, Marvin Eduardo'
              or [GI First Name] is null
          );

    update dbo.[TB Intakes]
    set [GI Last Name] = N'Carpio',
        [GI First Name] = N'Alejandro'
    where [ID] = 193
          and
          (
              [GI Last Name] is null
              or [GI First Name] is null
          );

    update dbo.[TB Intakes]
    set [GI Last Name] = N'Kerr'
    where [ID] = 2482
          and [GI Last Name] is null;

    commit;
    print N'    SECTION 4 done.';
end try
begin catch
    if xact_state() <> 0
        rollback;
    throw;
end catch;
go


-- #############################################################################
-- SECTION 5 — PHASE 1b: path migration (S:\ → /Company/)
--   Rewrites tblCaseDocuments.DocumentFileName, tblScans.ScanLocation, and
--   [TB Intakes].[Scan Location GI] across 11 per-category passes. Rows the
--   rewrite cannot recover are quarantined to the manual-triage tables from
--   SECTION 2.
--
--   AUTO-COMMIT POLICY: commits only if the leftover-offender count = 0
--   across all three tables (excluding manual-triage rows). On any error or
--   non-zero offender count: ROLLBACK + THROW.
-- #############################################################################

print N'';
print N'>>> SECTION 5: Phase 1b — path migration';
go

begin try
    begin transaction MigratePathsToDropbox;

    -- -----------------------------------------------------------------------
    -- PART A — tblCaseDocuments.DocumentFileName (single 'S:\…' / '#S:\…#' pass)
    -- -----------------------------------------------------------------------
    declare @CountDocsBefore int,
            @CountDocsUpdated int;

    select @CountDocsBefore = count(*)
    from dbo.tblCaseDocuments
    where DocumentFileName is not null
          and DocumentFileName <> '';

    update dbo.tblCaseDocuments
    set DocumentFileName = replace(replace(   case
                                                  when left(DocumentFileName, 1) = '#'
                                                       and right(DocumentFileName, 1) = '#' then
                                                      substring(DocumentFileName, 2, len(DocumentFileName) - 2)
                                                  when left(DocumentFileName, 1) = '#' then
                                                      substring(DocumentFileName, 2, len(DocumentFileName) - 1)
                                                  when right(DocumentFileName, 1) = '#' then
                                                      left(DocumentFileName, len(DocumentFileName) - 1)
                                                  else
                                                      DocumentFileName
                                              end,
                                              'S:\',
                                              '/Company/'
                                          ),
                                   '\',
                                   '/'
                                  )
    where DocumentFileName is not null
          and DocumentFileName <> ''
          and
          (
              DocumentFileName like 'S:\%'
              or DocumentFileName like '#S:\%'
          );
    set @CountDocsUpdated = @@rowcount;

    -- -----------------------------------------------------------------------
    -- PART B — tblScans.ScanLocation (six per-category passes + quarantine)
    -- -----------------------------------------------------------------------
    declare @CountScansBefore int,
            @CountScansB1 int = 0,
            @CountScansB2 int = 0,
            @CountScansB3 int = 0,
            @CountScansB4 int = 0,
            @CountScansB5 int = 0,
            @CountScansB6 int = 0,
            @CountScansQuarantined int = 0;

    select @CountScansBefore = count(*)
    from dbo.tblScans
    where ScanLocation is not null
          and ScanLocation <> '';

    -- B-1: '#S:\…#' (Access hyperlink wrapper around a bare S:\ path)
    update dbo.tblScans
    set ScanLocation = replace(
                                  replace(
                                             replace(
                                                        substring(
                                                                     ScanLocation,
                                                                     2,
                                                                     len(ScanLocation)
                                                                     - case
                                                                           when right(ScanLocation, 1) = '#' then
                                                                               2
                                                                           else
                                                                               1
                                                                       end
                                                                 ),
                                                        '%20',
                                                        ' '
                                                    ),
                                             'S:\',
                                             '/Company/'
                                         ),
                                  '\',
                                  '/'
                              )
    where ScanLocation like '#S:\%';
    set @CountScansB1 = @@rowcount;

    -- B-2: 'S:\…path#\\TBF-SRVR12\…#' (displaytext + hyperlink-URL suffix)
    update dbo.tblScans
    set ScanLocation = replace(replace(replace(   case
                                                      when charindex('#', ScanLocation) > 0 then
                                                          left(ScanLocation, charindex('#', ScanLocation) - 1)
                                                      else
                                                          ScanLocation
                                                  end,
                                                  '%20',
                                                  ' '
                                              ),
                                       'S:\',
                                       '/Company/'
                                      ),
                               '\',
                               '/'
                              )
    where ScanLocation like 'S:\%';
    set @CountScansB2 = @@rowcount;

    -- B-3: '#file:///S:\…#' (URL-encoded with file:// wrapper)
    update dbo.tblScans
    set ScanLocation = replace(
                                  replace(
                                             replace(
                                                        substring(
                                                                     ScanLocation,
                                                                     2 + 8,
                                                                     len(ScanLocation)
                                                                     - case
                                                                           when right(ScanLocation, 1) = '#' then
                                                                               2
                                                                           else
                                                                               1
                                                                       end - 8
                                                                 ),
                                                        '%20',
                                                        ' '
                                                    ),
                                             'S:\',
                                             '/Company/'
                                         ),
                                  '\',
                                  '/'
                              )
    where ScanLocation like '#file:///S:\%';
    set @CountScansB3 = @@rowcount;

    -- B-4: '#?S:\…#' ('#?' typo prefix)
    update dbo.tblScans
    set ScanLocation = replace(
                                  replace(
                                             substring(
                                                          ScanLocation,
                                                          3,
                                                          len(ScanLocation)
                                                          - case
                                                                when right(ScanLocation, 1) = '#' then
                                                                    3
                                                                else
                                                                    2
                                                            end
                                                      ),
                                             'S:\',
                                             '/Company/'
                                         ),
                                  '\',
                                  '/'
                              )
    where ScanLocation like '#?S:\%';
    set @CountScansB4 = @@rowcount;

    -- B-5: '#file:///\\TBF-SRVR12\<co>\…#' (legacy UNC + file:// wrapper)
    update dbo.tblScans
    set ScanLocation = replace(
                                  replace(
                                             replace(
                                                        replace(
                                                                   substring(
                                                                                ScanLocation,
                                                                                2 + 8,
                                                                                len(ScanLocation)
                                                                                - case
                                                                                      when right(ScanLocation, 1) = '#' then
                                                                                          2
                                                                                      else
                                                                                          1
                                                                                  end - 8
                                                                            ),
                                                                   '%20',
                                                                   ' '
                                                               ),
                                                        '\\TBF-SRVR12\Company\',
                                                        '/Company/'
                                                    ),
                                             'S:\',
                                             '/Company/'
                                         ),
                                  '\',
                                  '/'
                              )
    where ScanLocation like '#file:///\\TBF-SRVR12\%';
    set @CountScansB5 = @@rowcount;

    -- B-6: '#\\TBF-SRVR12\<co>\…#' (legacy UNC, bare)
    update dbo.tblScans
    set ScanLocation = replace(
                                  replace(
                                             replace(
                                                        substring(
                                                                     ScanLocation,
                                                                     2,
                                                                     len(ScanLocation)
                                                                     - case
                                                                           when right(ScanLocation, 1) = '#' then
                                                                               2
                                                                           else
                                                                               1
                                                                       end
                                                                 ),
                                                        '\\TBF-SRVR12\Company\',
                                                        '/Company/'
                                                    ),
                                             'S:\',
                                             '/Company/'
                                         ),
                                  '\',
                                  '/'
                              )
    where ScanLocation like '#\\TBF-SRVR12\%';
    set @CountScansB6 = @@rowcount;

    -- B-7: Quarantine — anything non-null/non-empty that matches none of the
    -- above and is not already migrated. Dedup by ScansID.
    insert into dbo.tblScans_ManualTriage
    (
        ScansID,
        OriginalValue,
        Reason
    )
    select s.ScansID,
           s.ScanLocation,
           N'Hash-less or mid-string-corrupted ScanLocation — no automatic rewrite'
    from dbo.tblScans s
    where s.ScanLocation is not null
          and ltrim(rtrim(s.ScanLocation)) <> ''
          and s.ScanLocation not like '/Company/%'
          and s.ScanLocation not like '#S:\%'
          and s.ScanLocation not like 'S:\%'
          and s.ScanLocation not like '#file:///S:\%'
          and s.ScanLocation not like '#?S:\%'
          and s.ScanLocation not like '#file:///\\TBF-SRVR12\%'
          and s.ScanLocation not like '#\\TBF-SRVR12\%'
          and not exists
                  (
                      select 1 from dbo.tblScans_ManualTriage q where q.ScansID = s.ScansID
                  );
    set @CountScansQuarantined = @@rowcount;

    -- -----------------------------------------------------------------------
    -- PART C — [TB Intakes].[Scan Location GI] (four passes + quarantine)
    -- -----------------------------------------------------------------------
    declare @CountIntakesBefore int,
            @CountIntakesC1 int = 0,
            @CountIntakesC2 int = 0,
            @CountIntakesC3 int = 0,
            @CountIntakesC4 int = 0,
            @CountIntakesQuarantined int = 0;

    select @CountIntakesBefore = count(*)
    from [TB Intakes]
    where [Scan Location GI] is not null
          and [Scan Location GI] <> '';

    -- C-1: 'S:\…' (bare path)
    update [TB Intakes]
    set [Scan Location GI] = replace(replace([Scan Location GI], 'S:\', '/Company/'), '\', '/')
    where [Scan Location GI] like 'S:\%';
    set @CountIntakesC1 = @@rowcount;

    -- C-2: '#S:\…#'
    update [TB Intakes]
    set [Scan Location GI] = replace(
                                        replace(
                                                   substring(
                                                                [Scan Location GI],
                                                                2,
                                                                len([Scan Location GI])
                                                                - case
                                                                      when right([Scan Location GI], 1) = '#' then
                                                                          2
                                                                      else
                                                                          1
                                                                  end
                                                            ),
                                                   'S:\',
                                                   '/Company/'
                                               ),
                                        '\',
                                        '/'
                                    )
    where [Scan Location GI] like '#S:\%';
    set @CountIntakesC2 = @@rowcount;

    -- C-3: '#file:///S:\…#'
    update [TB Intakes]
    set [Scan Location GI] = replace(
                                        replace(
                                                   replace(
                                                              substring(
                                                                           [Scan Location GI],
                                                                           2 + 8,
                                                                           len([Scan Location GI])
                                                                           - case
                                                                                 when right([Scan Location GI], 1) = '#' then
                                                                                     2
                                                                                 else
                                                                                     1
                                                                             end - 8
                                                                       ),
                                                              '%20',
                                                              ' '
                                                          ),
                                                   'S:\',
                                                   '/Company/'
                                               ),
                                        '\',
                                        '/'
                                    )
    where [Scan Location GI] like '#file:///S:\%';
    set @CountIntakesC3 = @@rowcount;

    -- C-4: '?S:\…'
    update [TB Intakes]
    set [Scan Location GI] = replace(
                                        replace(
                                                   substring([Scan Location GI], 2, len([Scan Location GI]) - 1),
                                                   'S:\',
                                                   '/Company/'
                                               ),
                                        '\',
                                        '/'
                                    )
    where [Scan Location GI] like '?S:\%';
    set @CountIntakesC4 = @@rowcount;

    -- C-5: Quarantine intake rows that match no pattern.
    insert into dbo.TBIntakes_ManualTriage
    (
        IntakeID,
        GILastName,
        GIFirstName,
        GIDate,
        OriginalValue,
        Reason
    )
    select i.[ID],
           i.[GI Last Name],
           i.[GI First Name],
           i.[GI Date],
           i.[Scan Location GI],
           N'Hash-less or root-missing Scan Location GI — no automatic rewrite'
    from [TB Intakes] i
    where i.[Scan Location GI] is not null
          and ltrim(rtrim(i.[Scan Location GI])) <> ''
          and i.[Scan Location GI] not like '/Company/%'
          and i.[Scan Location GI] not like 'S:\%'
          and i.[Scan Location GI] not like '#S:\%'
          and i.[Scan Location GI] not like '#file:///S:\%'
          and i.[Scan Location GI] not like '?S:\%'
          and not exists
                  (
                      select 1 from dbo.TBIntakes_ManualTriage q where q.IntakeID = i.[ID]
                  );
    set @CountIntakesQuarantined = @@rowcount;

    -- -----------------------------------------------------------------------
    -- Leftover-offender check inside the transaction — bail out if non-zero.
    -- Quarantined rows are excluded; they're intentionally untouched.
    -- -----------------------------------------------------------------------
    declare @LeftoverDocs int,
            @LeftoverScans int,
            @LeftoverIntakes int;

    select @LeftoverDocs = count(*)
    from dbo.tblCaseDocuments
    where DocumentFileName is not null
          and DocumentFileName <> ''
          and
          (
              DocumentFileName like '%\%'
              or DocumentFileName like '%S:%'
              or left(DocumentFileName, 1) = '#'
              or right(DocumentFileName, 1) = '#'
              or DocumentFileName like '%file:///%'
              or DocumentFileName like '%\\TBF-SRVR12\%'
              or DocumentFileName like '%[%]20%' collate Latin1_General_BIN
          );

    select @LeftoverScans = count(*)
    from dbo.tblScans s
    where s.ScanLocation is not null
          and s.ScanLocation <> ''
          and not exists
                  (
                      select 1 from dbo.tblScans_ManualTriage q where q.ScansID = s.ScansID
                  )
          and
          (
              s.ScanLocation like '%\%'
              or s.ScanLocation like '%S:%'
              or left(s.ScanLocation, 1) = '#'
              or right(s.ScanLocation, 1) = '#'
              or s.ScanLocation like '%file:///%'
              or s.ScanLocation like '%\\TBF-SRVR12\%'
              or s.ScanLocation like '%[%]20%' collate Latin1_General_BIN
          );

    select @LeftoverIntakes = count(*)
    from [TB Intakes]
    where [Scan Location GI] is not null
          and [Scan Location GI] <> ''
          and not exists
                  (
                      select 1
                      from dbo.TBIntakes_ManualTriage q
                      where q.OriginalValue = [TB Intakes].[Scan Location GI]
                  )
          and
          (
              [Scan Location GI] like '%\%'
              or [Scan Location GI] like '%S:%'
              or left([Scan Location GI], 1) = '#'
              or right([Scan Location GI], 1) = '#'
              or [Scan Location GI] like '%file:///%'
              or [Scan Location GI] like '%\\TBF-SRVR12\%'
              or [Scan Location GI] like '%[%]20%' collate Latin1_General_BIN
          );

    if @LeftoverDocs + @LeftoverScans + @LeftoverIntakes > 0
    begin
        rollback transaction MigratePathsToDropbox;
        declare @msg varchar(400)
            = N'SECTION 5 aborted — leftover offenders detected. ' + N'tblCaseDocuments='
              + cast(@LeftoverDocs as varchar(10)) + N' tblScans=' + cast(@LeftoverScans as varchar(10))
              + N' [TB Intakes]=' + cast(@LeftoverIntakes as varchar(10));
        throw 51002, @msg, 1;
    end;

    commit transaction MigratePathsToDropbox;

    print N'    SECTION 5 done.';
    print N'    Per-pass counts:';
    print N'      tblCaseDocuments: before=' + cast(@CountDocsBefore as varchar(10)) + N' updated='
          + cast(@CountDocsUpdated as varchar(10));
    print N'      tblScans:         before=' + cast(@CountScansBefore as varchar(10)) + N' B1='
          + cast(@CountScansB1 as varchar(10)) + N' B2=' + cast(@CountScansB2 as varchar(10)) + N' B3='
          + cast(@CountScansB3 as varchar(10)) + N' B4=' + cast(@CountScansB4 as varchar(10)) + N' B5='
          + cast(@CountScansB5 as varchar(10)) + N' B6=' + cast(@CountScansB6 as varchar(10)) + N' quarantined='
          + cast(@CountScansQuarantined as varchar(10));
    print N'      [TB Intakes]:     before=' + cast(@CountIntakesBefore as varchar(10)) + N' C1='
          + cast(@CountIntakesC1 as varchar(10)) + N' C2=' + cast(@CountIntakesC2 as varchar(10)) + N' C3='
          + cast(@CountIntakesC3 as varchar(10)) + N' C4=' + cast(@CountIntakesC4 as varchar(10)) + N' quarantined='
          + cast(@CountIntakesQuarantined as varchar(10));
end try
begin catch
    if xact_state() <> 0
        rollback transaction MigratePathsToDropbox;
    throw;
end catch;
go


-- #############################################################################
-- SECTION 6 — VERIFICATION (post-migration sanity)
--   Expected: every numeric column in the leftover-offender block reads 0.
-- #############################################################################

print N'';
print N'>>> SECTION 6: verification — leftover offenders (expect all zeros)';
go

select 'tblCaseDocuments' as TableName,
       sum(   case
                  when DocumentFileName like '%\%' then
                      1
                  else
                      0
              end
          ) as HasBackslash,
       sum(   case
                  when DocumentFileName like '%S:%' then
                      1
                  else
                      0
              end
          ) as HasSColon,
       sum(   case
                  when left(DocumentFileName, 1) = '#' then
                      1
                  else
                      0
              end
          ) as LeadingHash,
       sum(   case
                  when right(DocumentFileName, 1) = '#' then
                      1
                  else
                      0
              end
          ) as TrailingHash,
       sum(   case
                  when DocumentFileName like '%file:///%' then
                      1
                  else
                      0
              end
          ) as HasFileURL,
       sum(   case
                  when DocumentFileName like '%\\TBF-SRVR12\%' then
                      1
                  else
                      0
              end
          ) as HasLegacyUNC,
       sum(   case
                  when DocumentFileName like '%[%]20%' collate Latin1_General_BIN then
                      1
                  else
                      0
              end
          ) as HasUrlEncodedSpace
from dbo.tblCaseDocuments
where DocumentFileName is not null
      and DocumentFileName <> '';

select 'tblScans' as TableName,
       sum(   case
                  when s.ScanLocation like '%\%' then
                      1
                  else
                      0
              end
          ) as HasBackslash,
       sum(   case
                  when s.ScanLocation like '%S:%' then
                      1
                  else
                      0
              end
          ) as HasSColon,
       sum(   case
                  when left(s.ScanLocation, 1) = '#' then
                      1
                  else
                      0
              end
          ) as LeadingHash,
       sum(   case
                  when right(s.ScanLocation, 1) = '#' then
                      1
                  else
                      0
              end
          ) as TrailingHash,
       sum(   case
                  when s.ScanLocation like '%file:///%' then
                      1
                  else
                      0
              end
          ) as HasFileURL,
       sum(   case
                  when s.ScanLocation like '%\\TBF-SRVR12\%' then
                      1
                  else
                      0
              end
          ) as HasLegacyUNC,
       sum(   case
                  when s.ScanLocation like '%[%]20%' collate Latin1_General_BIN then
                      1
                  else
                      0
              end
          ) as HasUrlEncodedSpace
from dbo.tblScans s
where s.ScanLocation is not null
      and s.ScanLocation <> ''
      and not exists
              (
                  select 1 from dbo.tblScans_ManualTriage q where q.ScansID = s.ScansID
              );

select 'TB Intakes' as TableName,
       sum(   case
                  when [Scan Location GI] like '%\%' then
                      1
                  else
                      0
              end
          ) as HasBackslash,
       sum(   case
                  when [Scan Location GI] like '%S:%' then
                      1
                  else
                      0
              end
          ) as HasSColon,
       sum(   case
                  when left([Scan Location GI], 1) = '#' then
                      1
                  else
                      0
              end
          ) as LeadingHash,
       sum(   case
                  when right([Scan Location GI], 1) = '#' then
                      1
                  else
                      0
              end
          ) as TrailingHash,
       sum(   case
                  when [Scan Location GI] like '%file:///%' then
                      1
                  else
                      0
              end
          ) as HasFileURL,
       sum(   case
                  when [Scan Location GI] like '%\\TBF-SRVR12\%' then
                      1
                  else
                      0
              end
          ) as HasLegacyUNC,
       sum(   case
                  when [Scan Location GI] like '%[%]20%' collate Latin1_General_BIN then
                      1
                  else
                      0
              end
          ) as HasUrlEncodedSpace
from [TB Intakes]
where [Scan Location GI] is not null
      and [Scan Location GI] <> ''
      and not exists
              (
                  select 1
                  from dbo.TBIntakes_ManualTriage q
                  where q.OriginalValue = [TB Intakes].[Scan Location GI]
              );
go

print N'';
print N'>>> SECTION 6: verification — path-prefix distribution';
go

select 'tblCaseDocuments' as TableName,
       left(DocumentFileName, 30) as PathPrefix,
       count(*) as [RowCount]
from dbo.tblCaseDocuments
where DocumentFileName is not null
      and DocumentFileName <> ''
group by left(DocumentFileName, 30)
order by count(*) desc;

select 'tblScans' as TableName,
       left(s.ScanLocation, 30) as PathPrefix,
       count(*) as [RowCount]
from dbo.tblScans s
where s.ScanLocation is not null
      and s.ScanLocation <> ''
      and not exists
              (
                  select 1 from dbo.tblScans_ManualTriage q where q.ScansID = s.ScansID
              )
group by left(s.ScanLocation, 30)
order by count(*) desc;

select 'TB Intakes' as TableName,
       left([Scan Location GI], 30) as PathPrefix,
       count(*) as [RowCount]
from [TB Intakes]
where [Scan Location GI] is not null
      and [Scan Location GI] <> ''
      and not exists
              (
                  select 1
                  from dbo.TBIntakes_ManualTriage q
                  where q.OriginalValue = [TB Intakes].[Scan Location GI]
              )
group by left([Scan Location GI], 30)
order by count(*) desc;
go


-- #############################################################################
-- SECTION 7 — PHASE 1b LISTINGS (output only — no auto-fix)
--   These items still require per-row human triage. The script surfaces what
--   needs attention; the IT admin authors per-row fixes and re-runs the
--   script (or appends fixes inline).
-- #############################################################################

print N'';
print N'>>> SECTION 7.1: Phase 1b B-9 — path-length pre-flight (paths > 260 chars)';
go

-- Dropbox enforces ~260 chars effective path length (G14). Listing any post-
-- rewrite path that exceeds this limit. Expected (test DB, 2026-05-14): zero
-- rows over 260 chars; max length is 247 chars in tblCaseDocuments.
select 'tblCaseDocuments' as TableName,
       CaseDocumentID,
       len(DocumentFileName) as PathLength,
       DocumentFileName
from dbo.tblCaseDocuments
where DocumentFileName is not null
      and len(DocumentFileName) > 260
order by len(DocumentFileName) desc;

select 'tblScans' as TableName,
       ScansID,
       len(ScanLocation) as PathLength,
       ScanLocation
from dbo.tblScans s
where s.ScanLocation is not null
      and len(s.ScanLocation) > 260
      and not exists
              (
                  select 1 from dbo.tblScans_ManualTriage q where q.ScansID = s.ScansID
              )
order by len(ScanLocation) desc;

select 'TB Intakes' as TableName,
       [ID] as IntakeID,
       len([Scan Location GI]) as PathLength,
       [Scan Location GI]
from [TB Intakes]
where [Scan Location GI] is not null
      and len([Scan Location GI]) > 260
order by len([Scan Location GI]) desc;
go

-- Distribution view: how many rows fall in each length bucket (informational).
select 'tblCaseDocuments' as TableName,
       sum(   case
                  when len(DocumentFileName)
                       between 1 and 100 then
                      1
                  else
                      0
              end
          ) as Bucket_1_100,
       sum(   case
                  when len(DocumentFileName)
                       between 101 and 150 then
                      1
                  else
                      0
              end
          ) as Bucket_101_150,
       sum(   case
                  when len(DocumentFileName)
                       between 151 and 200 then
                      1
                  else
                      0
              end
          ) as Bucket_151_200,
       sum(   case
                  when len(DocumentFileName)
                       between 201 and 260 then
                      1
                  else
                      0
              end
          ) as Bucket_201_260,
       sum(   case
                  when len(DocumentFileName) > 260 then
                      1
                  else
                      0
              end
          ) as Bucket_Over260,
       max(len(DocumentFileName)) as MaxLen
from dbo.tblCaseDocuments
where DocumentFileName is not null
      and DocumentFileName <> '';
go

print N'';
print N'>>> SECTION 7.2: Phase 1b B-4 — 13 bracket-literal rows in tblCaseDocuments';
go

-- Classified into two defect classes — the 9 unresolved-template rows can be
-- re-resolved by re-running spGetDocumentFolderName + spGetDocumentFileName
-- after the vwfrmClientLedger row is repopulated. The 4 truncated rows ending
-- with '[df' need the on-disk file inspected to recover the intended filename.
-- LIKE pattern notes: [[]  = literal '['; [_] = literal '_' (escape the
-- wildcard); trailing ']' is literal outside an active bracket class.
select CaseDocumentID,
       CaseID,
       DocumentFileName,
       case
           when DocumentFileName like '%[[]Case[_]Letter]%' then
               'unresolved template (re-resolve via SPs)'
           when DocumentFileName like '%[[]df' then
               'truncated filename ending [df (inspect on-disk file)'
           else
               'other bracket defect (manual triage)'
       end as DefectType
from dbo.tblCaseDocuments
where DocumentFileName like '%[[]%'
order by case
             when DocumentFileName like '%[[]Case[_]Letter]%' then
                 1
             when DocumentFileName like '%[[]df' then
                 2
             else
                 3
         end,
         CaseDocumentID;
go

print N'';
print N'>>> SECTION 7.3: Phase 1b B-6 — non-canonical roots in tblCaseDocuments';
go

-- Rows whose root prefix does NOT match the canonical /Company/COMMON/<Atty>/_CLIENTS/...
-- pattern. Per the plan, every row must be remediated to a verified Dropbox path
-- before Phase 7 — "skip" is not an option. Grouped by 4-segment prefix.
with NonCanonical
as (select CaseDocumentID,
           CaseID,
           DocumentFileName,
           -- Extract the first 4 path segments after /Company/ for grouping
           left(DocumentFileName, 60) as PathPrefix60
    from dbo.tblCaseDocuments
    where DocumentFileName is not null
          and DocumentFileName <> ''
          and DocumentFileName like '/Company/%'
          and DocumentFileName not like '/Company/COMMON/%/_CLIENTS/%'
          and DocumentFileName not like '/Company/Closed File Scans/%')
select PathPrefix60,
       count(*) as [RowCount]
from NonCanonical
group by PathPrefix60
order by count(*) desc;

-- Sample rows per non-canonical prefix (first 30 rows overall — review in SSMS)
select top 30
       CaseDocumentID,
       CaseID,
       DocumentFileName
from dbo.tblCaseDocuments
where DocumentFileName is not null
      and DocumentFileName <> ''
      and DocumentFileName like '/Company/%'
      and DocumentFileName not like '/Company/COMMON/%/_CLIENTS/%'
      and DocumentFileName not like '/Company/Closed File Scans/%'
order by DocumentFileName;
go

print N'';
print N'>>> SECTION 7.4: Phase 1b B-7 — tblScans manual-triage queue';
go

select TriageID = row_number() over (order by ScansID),
       ScansID,
       OriginalValue,
       Reason,
       QuarantinedAt
from dbo.tblScans_ManualTriage
order by ScansID;
go

print N'';
print N'>>> SECTION 7.5: Phase 1b B-8 — [TB Intakes] manual-triage queue';
go

select TriageID,
       IntakeID,
       GILastName,
       GIFirstName,
       GIDate,
       OriginalValue,
       Reason,
       QuarantinedAt
from dbo.TBIntakes_ManualTriage
order by TriageID;
go


-- #############################################################################
-- POST-INSTALL CONFIRMATION
-- #############################################################################

print N'';
print N'>>> Post-install confirmation';
go

select name,
       type_desc
from sys.objects
where name in ( N'tblDropboxRootConfig', N'tblDropboxConfig', N'tblDropboxRevocationList', N'tblDropboxAuditLog',
                N'tblDropboxOrphanQueue', N'tblDropboxVerificationReport', N'spLogDropboxAuditEvent',
                N'tblScans_ManualTriage', N'TBIntakes_ManualTriage'
              )
order by name;

select ConfigID,
       NamespaceId,
       TeamRootPath
from dbo.tblDropboxRootConfig;

select ConfigID,
       AppKey,
       case
           when AppSecret = N'REPLACE_WITH_APP_SECRET_FROM_DROPBOX_APP_CONSOLE' then
               '<<PLACEHOLDER — IT MUST UPDATE>>'
           else
               '<set>'
       end as AppSecretStatus,
       RedirectUri
from dbo.tblDropboxConfig;
go

print N'';
print N'================================================================================';
print N'INSTALL_all.sql — complete.';
print N'  Don''t forget: UPDATE dbo.tblDropboxConfig SET AppSecret = N''<real>'' WHERE ConfigID = 1;';
print N'  Review SECTION 7 output for B-4, B-6, B-7, B-8, B-9 triage decisions.';
print N'================================================================================';
go
