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
--   SECTION 8 — Phase 4c SP rewrites: drop+recreate the 7 stored procedures
--                                     that resolved S:\-rooted paths against
--                                     tblDocumentRootDirectory, replacing them
--                                     with bodies that resolve /Company/-
--                                     rooted Dropbox paths against
--                                     tblDropboxRootConfig. Includes the G2
--                                     spMoveDocumentFolder rewrite and the
--                                     G13 spSaveCaseDocument token-validation
--                                     guard. Legacy bodies are preserved as
--                                     commented-out blocks immediately above
--                                     each rewrite for single-file rollback.
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

-- #############################################################################
-- SECTION 8 — PHASE 4c: STORED PROCEDURE REWRITES (Dropbox path-aware)
-- #############################################################################
--
-- Drops + recreates the 7 stored procedures listed in plan section "Updated
-- stored procedures":
--   8.1 spGetIntakeFolderName         — read IntakeDirectory from
--                                       tblDropboxRootConfig
--   8.2 spGetDocumentFolderName       — open-case folder resolver
--   8.3 spGetClosedDocumentFolderName — closed-case folder resolver
--   8.4 spGetClosedFileScanFolderName — closed-file-scan folder resolver
--   8.5 spGetAllInvoicesFolderName    — firm-wide invoices folder resolver
--   8.6 spMoveDocumentFolder          — G2 rewrite (new signature:
--                                       @OldFolderPath/@NewFolderPath,
--                                       SET XACT_ABORT ON, both-zero-rowcount
--                                       hard-fail, updates tblCaseDocuments
--                                       AND tblScans in one transaction)
--   8.7 spSaveCaseDocument            — G13 token-validation guard
--                                       (THROW on [Field]/<currentdate>/
--                                       (customuserentry) substrings in
--                                       @DocumentName)
--   8.8 Verification                  — exec each path-building SP for the
--                                       QUAIL/30337 known-good case and
--                                       print the resolved /Company/ path
--                                       for visual sanity check.
--
-- The four SPs that DON'T need rewriting in 4c (and are intentionally NOT
-- touched here):
--   spGetCaseDocument           — already returns tblCaseDocuments.DocumentFileName
--                                 verbatim (now /Company/-rooted post-migration).
--   spGetDocumentFileName       — produces a FILENAME from DocumentNamingRule
--                                 (no path separator concerns).
--   spGetIntakeDocumentFileName — same, for intake filename.
--   fnGetListOfWords            — tokenizer used by the path-building SPs;
--                                 preserved verbatim (works for both / and \).
--
-- Rollback safety: the legacy (pre-4c) body of every rewritten SP is
-- preserved as a commented-out block immediately above its replacement.
-- To roll back one SP: comment out the new create-procedure, uncomment the
-- legacy create-procedure, re-run this installer (idempotent).
--
-- Path-separator handling: tblDropboxRootConfig still stores the legacy
-- template strings with `\` separators (e.g.,
-- '\ [Orig_Atty] \_CLIENTS\ [Case_Letter] \ ...') so the tokenizer is
-- unchanged. Each rewritten SP wraps its dynamic-SQL output in
-- REPLACE(..., '\', '/') so the final resolved path is forward-slash
-- form. This avoids touching the template-storage format itself.

print N'';
print N'>>> SECTION 8: Phase 4c — Dropbox-aware stored procedure rewrites';
go


-- ---------------------------------------------------------------------------
-- 8.1 spGetIntakeFolderName — single-row read of IntakeDirectory from
--     tblDropboxRootConfig (replaces the legacy read against
--     tblDocumentRootDirectory).
-- ---------------------------------------------------------------------------

print N'    8.1 dropping + recreating spGetIntakeFolderName';
go

-- LEGACY (pre-Phase 4c) — preserved for rollback. To roll back: comment out
-- the new create-procedure below, uncomment the legacy version, re-run.
-- ============================================================================
-- CREATE PROCEDURE [dbo].[spGetIntakeFolderName]
-- AS
-- BEGIN
--     /*
--         exec spGetIntakeFolderName
--     */
--     SELECT IntakeDirectory AS DocumentFolder
--     FROM dbo.tblDocumentRootDirectory (NOLOCK);
-- END;
-- ============================================================================

drop procedure if exists dbo.spGetIntakeFolderName;
go

create procedure dbo.spGetIntakeFolderName
as
begin
    set nocount on;

    select IntakeDirectory as DocumentFolder
    from dbo.tblDropboxRootConfig with (nolock)
    where ConfigID = 1;
end;
go


-- ---------------------------------------------------------------------------
-- 8.2 spGetDocumentFolderName — open-case folder resolver. Builds the path
--     by tokenizing tblDropboxRootConfig.DocumentRootNaming and substituting
--     contract columns from vwfrmClientLedger (with the Case_Letter CodeVal
--     lookup via tblDropD). Wraps final output in REPLACE('\','/') to emit
--     a Dropbox forward-slash path.
-- ---------------------------------------------------------------------------

print N'    8.2 dropping + recreating spGetDocumentFolderName';
go

-- LEGACY (pre-Phase 4c) — preserved for rollback.
-- ============================================================================
-- CREATE   PROCEDURE [dbo].[spGetDocumentFolderName]
--     @DocumentType VARCHAR(250),
--     @CaseID INT
-- AS
-- BEGIN
--     /*
--         exec spGetDocumentFolderName 'Client Documents', 26081
--     */
--     DECLARE @DocumentRootNaming VARCHAR(500);
--     DECLARE @DocumentRootDirectory VARCHAR(500);
--     DECLARE @DocumentFolder VARCHAR(500);
--     DECLARE @sql VARCHAR(MAX);
--     DECLARE @word VARCHAR(500);
--
--     SELECT @DocumentRootNaming = DocumentRootNaming,
--            @DocumentRootDirectory = DocumentRootDirectory
--     FROM dbo.tblDocumentRootDirectory (NOLOCK);
--
--     SELECT @DocumentFolder = DocumentFolder
--     FROM dbo.tblDocumentTypes (NOLOCK)
--     WHERE DocumentType = @DocumentType;
--
--     SELECT @sql = 'SELECT LTRIM(RTRIM(''' + @DocumentRootDirectory + ''' +';
--
--     DECLARE cursorT CURSOR READ_ONLY FOR
--     SELECT Word
--     FROM dbo.fnGetListOfWords(@DocumentRootNaming, ' ')
--     ORDER BY Position;
--
--     OPEN cursorT;
--     WHILE (1 = 1)
--     BEGIN
--         FETCH cursorT INTO @word;
--         IF @@FETCH_STATUS <> 0 BREAK;
--         IF LEFT(@word, 1) = '['
--         BEGIN
--             SELECT @sql = @sql + ' convert(varchar(250), isnull(' + @word + ', '''+ '' + '''))+';
--         END;
--         ELSE IF @word = '~'
--         BEGIN
--             SELECT @sql = @sql + '''' + ' ' + ''' +';
--         END;
--         ELSE
--         BEGIN
--             SELECT @sql = @sql + '''' + @word + ''' +';
--         END;
--     END;
--
--     SELECT @sql = LEFT(@sql, LEN(@sql) - 1) + ' + ''' + @DocumentFolder + ''')) AS DocumentFolder ';
--     SELECT @sql = @sql + ' FROM (select c.CaseID, c.Orig_Atty, d.CodeVal as Case_Letter,
--                                         c.Last_Name, c.First_Name, c.FileNo
--                                         from vwfrmClientLedger c (nolock)
--                                         inner join tblDropD d (nolock)
--                                         on c.Case_Letter = d.Code
--                                         where d.FieldName = ''Case_Letter'') as X
--                                         WHERE CaseID = ' + CONVERT(VARCHAR(10), @CaseID);
--     EXEC (@sql);
--     CLOSE cursorT;
--     DEALLOCATE cursorT;
-- END;
-- ============================================================================

drop procedure if exists dbo.spGetDocumentFolderName;
go

create procedure dbo.spGetDocumentFolderName
    @DocumentType varchar(250),
    @CaseID int
as
begin
    set nocount on;
    /*
        exec dbo.spGetDocumentFolderName 'General', 30337
    */
    declare @DocumentRootNaming varchar(500);
    declare @TeamRootPath varchar(500);
    declare @DocumentFolder varchar(500);
    declare @sql varchar(max);
    declare @word varchar(500);

    select @DocumentRootNaming = DocumentRootNaming,
           @TeamRootPath = TeamRootPath
    from dbo.tblDropboxRootConfig with (nolock)
    where ConfigID = 1;

    select @DocumentFolder = DocumentFolder
    from dbo.tblDocumentTypes with (nolock)
    where DocumentType = @DocumentType;

    select @sql = 'SELECT REPLACE(LTRIM(RTRIM(''' + @TeamRootPath + ''' +';

    declare cursorT cursor read_only for
    select Word
    from dbo.fnGetListOfWords(@DocumentRootNaming, ' ')
    order by Position;

    open cursorT;
    while (1 = 1)
    begin
        fetch cursorT into @word;
        if @@fetch_status <> 0 break;

        if left(@word, 1) = '['
        begin
            select @sql = @sql + ' convert(varchar(250), isnull(' + @word + ', '''+ '' + '''))+';
        end;
        else if @word = '~'
        begin
            select @sql = @sql + '''' + ' ' + ''' +';
        end;
        else
        begin
            select @sql = @sql + '''' + @word + ''' +';
        end;
    end;

    select @sql = left(@sql, len(@sql) - 1) + ' + ''' + @DocumentFolder + ''')), ''\'', ''/'') AS DocumentFolder ';

    select @sql = @sql
        + ' FROM (select c.CaseID, c.Orig_Atty, d.CodeVal as Case_Letter,
                                    c.Last_Name, c.First_Name, c.FileNo
                                    from vwfrmClientLedger c (nolock)
                                    inner join tblDropD d (nolock)
                                    on c.Case_Letter = d.Code
                                    where d.FieldName = ''Case_Letter'') as X
                                    WHERE CaseID = ' + convert(varchar(10), @CaseID);

    exec (@sql);

    close cursorT;
    deallocate cursorT;
end;
go


-- ---------------------------------------------------------------------------
-- 8.3 spGetClosedDocumentFolderName — closed-case folder resolver. Same
--     pattern as 8.2 but uses tblDropboxRootConfig.DocumentClosedNaming.
-- ---------------------------------------------------------------------------

print N'    8.3 dropping + recreating spGetClosedDocumentFolderName';
go

-- LEGACY (pre-Phase 4c) — preserved for rollback.
-- ============================================================================
-- CREATE   PROCEDURE [dbo].[spGetClosedDocumentFolderName]
--     @DocumentType VARCHAR(250),
--     @CaseID INT
-- AS
-- BEGIN
--     /*
--         exec spGetClosedDocumentFolderName 'Client Documents', 26633
--     */
--     DECLARE @DocumentClosedNaming VARCHAR(500);
--     DECLARE @DocumentFolder VARCHAR(500);
--     DECLARE @DocumentRootDirectory VARCHAR(500);
--     DECLARE @sql VARCHAR(MAX);
--     DECLARE @word VARCHAR(500);
--
--     SELECT @DocumentClosedNaming = DocumentClosedNaming,
--            @DocumentRootDirectory = DocumentRootDirectory
--     FROM dbo.tblDocumentRootDirectory (NOLOCK);
--
--     SELECT @DocumentFolder = DocumentFolder
--     FROM dbo.tblDocumentTypes (NOLOCK)
--     WHERE DocumentType = @DocumentType;
--
--     SELECT @sql = 'SELECT LTRIM(RTRIM(''' + @DocumentRootDirectory + ''' +';
--     -- ... [tokenizer cursor; same shape as 8.2 LEGACY] ...
--     SELECT @sql = LEFT(@sql, LEN(@sql) - 1) + ' + ''' + @DocumentFolder + ''')) AS DocumentFolder ';
--     SELECT @sql = @sql + ' FROM (...derived view...) WHERE CaseID = ' + CONVERT(VARCHAR(10), @CaseID);
--     EXEC (@sql);
-- END;
-- ============================================================================

drop procedure if exists dbo.spGetClosedDocumentFolderName;
go

create procedure dbo.spGetClosedDocumentFolderName
    @DocumentType varchar(250),
    @CaseID int
as
begin
    set nocount on;
    /*
        exec dbo.spGetClosedDocumentFolderName 'General', 30337
    */
    declare @DocumentClosedNaming varchar(500);
    declare @DocumentFolder varchar(500);
    declare @TeamRootPath varchar(500);
    declare @sql varchar(max);
    declare @word varchar(500);

    select @DocumentClosedNaming = DocumentClosedNaming,
           @TeamRootPath = TeamRootPath
    from dbo.tblDropboxRootConfig with (nolock)
    where ConfigID = 1;

    select @DocumentFolder = DocumentFolder
    from dbo.tblDocumentTypes with (nolock)
    where DocumentType = @DocumentType;

    select @sql = 'SELECT REPLACE(LTRIM(RTRIM(''' + @TeamRootPath + ''' +';

    declare cursorT cursor read_only for
    select Word
    from dbo.fnGetListOfWords(@DocumentClosedNaming, ' ')
    order by Position;

    open cursorT;
    while (1 = 1)
    begin
        fetch cursorT into @word;
        if @@fetch_status <> 0 break;

        if left(@word, 1) = '['
        begin
            select @sql = @sql + ' convert(varchar(250), isnull(' + @word + ', '''+ '' + '''))+';
        end;
        else if @word = '~'
        begin
            select @sql = @sql + '''' + ' ' + ''' +';
        end;
        else
        begin
            select @sql = @sql + '''' + @word + ''' +';
        end;
    end;

    select @sql = left(@sql, len(@sql) - 1) + ' + ''' + @DocumentFolder + ''')), ''\'', ''/'') AS DocumentFolder ';

    select @sql = @sql
        + ' FROM (select c.CaseID, c.Orig_Atty, d.CodeVal as Case_Letter,
                                    c.Last_Name, c.First_Name, c.FileNo
                                    from vwfrmClientLedger c (nolock)
                                    inner join tblDropD d (nolock)
                                    on c.Case_Letter = d.Code
                                    where d.FieldName = ''Case_Letter'') as X
                                    WHERE CaseID = ' + convert(varchar(10), @CaseID);

    exec (@sql);

    close cursorT;
    deallocate cursorT;
end;
go


-- ---------------------------------------------------------------------------
-- 8.4 spGetClosedFileScanFolderName — closed-file-scans folder resolver.
--     Reads ClosedFileScanDirectory + ClosedFileScanNaming from
--     tblDropboxRootConfig. Includes Yr in the derived view (template uses
--     [Yr]). Same REPLACE('\','/') wrap.
-- ---------------------------------------------------------------------------

print N'    8.4 dropping + recreating spGetClosedFileScanFolderName';
go

-- LEGACY (pre-Phase 4c) — preserved for rollback.
-- ============================================================================
-- CREATE PROCEDURE [dbo].[spGetClosedFileScanFolderName]
--     @DocumentType VARCHAR(250),
--     @CaseID INT
-- AS
-- BEGIN
--     /*
--         exec spGetClosedFileScanFolderName 'General', 26633
--     */
--     DECLARE @DocumentClosedNaming VARCHAR(500);
--     DECLARE @DocumentFolder VARCHAR(500);
--     DECLARE @DocumentRootDirectory VARCHAR(500);
--     DECLARE @sql VARCHAR(MAX);
--     DECLARE @word VARCHAR(500);
--
--     SELECT @DocumentClosedNaming = ClosedFileScanNaming,
--            @DocumentRootDirectory = ClosedFileScanDirectory
--     FROM dbo.tblDocumentRootDirectory (NOLOCK);
--
--     SELECT @DocumentFolder = DocumentFolder
--     FROM dbo.tblDocumentTypes (NOLOCK)
--     WHERE DocumentType = @DocumentType;
--     -- ... [tokenizer cursor; same shape as 8.2 LEGACY] ...
--     SELECT @sql = LEFT(@sql, LEN(@sql) - 1) + ' + ''' + @DocumentFolder + ''')) AS DocumentFolder ';
--     SELECT @sql = @sql + ' FROM (select c.CaseID, c.Yr, c.Orig_Atty, d.CodeVal as Case_Letter,
--                                          c.Last_Name, c.First_Name, c.FileNo
--                                          ... ) as X
--                                    WHERE CaseID = ' + CONVERT(VARCHAR(10), @CaseID);
--     EXEC (@sql);
-- END;
-- ============================================================================

drop procedure if exists dbo.spGetClosedFileScanFolderName;
go

create procedure dbo.spGetClosedFileScanFolderName
    @DocumentType varchar(250),
    @CaseID int
as
begin
    set nocount on;
    /*
        exec dbo.spGetClosedFileScanFolderName 'General', 30337
    */
    declare @ClosedFileScanNaming varchar(500);
    declare @DocumentFolder varchar(500);
    declare @ClosedFileScanDirectory varchar(500);
    declare @sql varchar(max);
    declare @word varchar(500);

    select @ClosedFileScanNaming = ClosedFileScanNaming,
           @ClosedFileScanDirectory = ClosedFileScanDirectory
    from dbo.tblDropboxRootConfig with (nolock)
    where ConfigID = 1;

    select @DocumentFolder = DocumentFolder
    from dbo.tblDocumentTypes with (nolock)
    where DocumentType = @DocumentType;

    select @sql = 'SELECT REPLACE(LTRIM(RTRIM(''' + @ClosedFileScanDirectory + ''' +';

    declare cursorT cursor read_only for
    select Word
    from dbo.fnGetListOfWords(@ClosedFileScanNaming, ' ')
    order by Position;

    open cursorT;
    while (1 = 1)
    begin
        fetch cursorT into @word;
        if @@fetch_status <> 0 break;

        if left(@word, 1) = '['
        begin
            select @sql = @sql + ' convert(varchar(250), isnull(' + @word + ', '''+ '' + '''))+';
        end;
        else if @word = '~'
        begin
            select @sql = @sql + '''' + ' ' + ''' +';
        end;
        else
        begin
            select @sql = @sql + '''' + @word + ''' +';
        end;
    end;

    select @sql = left(@sql, len(@sql) - 1) + ' + ''' + @DocumentFolder + ''')), ''\'', ''/'') AS DocumentFolder ';

    select @sql = @sql
        + ' FROM (select c.CaseID, c.Yr, c.Orig_Atty, d.CodeVal as Case_Letter,
                                    c.Last_Name, c.First_Name, c.FileNo
                                    from vwfrmClientLedger c (nolock)
                                    inner join tblDropD d (nolock)
                                    on c.Case_Letter = d.Code
                                    where d.FieldName = ''Case_Letter'') as X
                                    WHERE CaseID = ' + convert(varchar(10), @CaseID);

    exec (@sql);

    close cursorT;
    deallocate cursorT;
end;
go


-- ---------------------------------------------------------------------------
-- 8.5 spGetAllInvoicesFolderName — firm-wide all-invoices folder resolver.
--     Hard-codes DocumentType='General' (legacy behavior); reads
--     AllInvoicesDirectory + AllInvoicesNaming from tblDropboxRootConfig.
-- ---------------------------------------------------------------------------

print N'    8.5 dropping + recreating spGetAllInvoicesFolderName';
go

-- LEGACY (pre-Phase 4c) — preserved for rollback.
-- ============================================================================
-- CREATE PROCEDURE [dbo].[spGetAllInvoicesFolderName]
--     @CaseID INT
-- AS
-- BEGIN
--     /*
--         exec spGetAllInvoicesFolderName 9966
--     */
--     DECLARE @AllInvoicesNaming VARCHAR(500);
--     DECLARE @DocumentFolder VARCHAR(500);
--     DECLARE @DocumentRootDirectory VARCHAR(500);
--     DECLARE @sql VARCHAR(MAX);
--     DECLARE @word VARCHAR(500);
--
--     SELECT @AllInvoicesNaming = AllInvoicesNaming,
--            @DocumentRootDirectory = AllInvoicesDirectory
--     FROM dbo.tblDocumentRootDirectory (NOLOCK);
--
--     SELECT @DocumentFolder = DocumentFolder
--     FROM dbo.tblDocumentTypes (NOLOCK)
--     WHERE DocumentType = 'General'
--     -- ... [tokenizer cursor; same shape as 8.2 LEGACY] ...
--     EXEC (@sql);
-- END;
-- ============================================================================

drop procedure if exists dbo.spGetAllInvoicesFolderName;
go

create procedure dbo.spGetAllInvoicesFolderName
    @CaseID int
as
begin
    set nocount on;
    /*
        exec dbo.spGetAllInvoicesFolderName 30337
    */
    declare @AllInvoicesNaming varchar(500);
    declare @DocumentFolder varchar(500);
    declare @AllInvoicesDirectory varchar(500);
    declare @sql varchar(max);
    declare @word varchar(500);

    select @AllInvoicesNaming = AllInvoicesNaming,
           @AllInvoicesDirectory = AllInvoicesDirectory
    from dbo.tblDropboxRootConfig with (nolock)
    where ConfigID = 1;

    select @DocumentFolder = DocumentFolder
    from dbo.tblDocumentTypes with (nolock)
    where DocumentType = 'General';

    -- AllInvoicesNaming is empty in current config; the cursor still runs
    -- (zero iterations) and the SQL falls through to the constant root path.
    select @sql = 'SELECT REPLACE(LTRIM(RTRIM(''' + @AllInvoicesDirectory + ''' +';

    declare cursorT cursor read_only for
    select Word
    from dbo.fnGetListOfWords(@AllInvoicesNaming, ' ')
    order by Position;

    open cursorT;
    while (1 = 1)
    begin
        fetch cursorT into @word;
        if @@fetch_status <> 0 break;

        if left(@word, 1) = '['
        begin
            select @sql = @sql + ' convert(varchar(250), isnull(' + @word + ', '''+ '' + '''))+';
        end;
        else if @word = '~'
        begin
            select @sql = @sql + '''' + ' ' + ''' +';
        end;
        else
        begin
            select @sql = @sql + '''' + @word + ''' +';
        end;
    end;

    -- If AllInvoicesNaming was empty, @sql ends with '''+' (no trailing +)
    -- and LEFT(...,LEN-1) would corrupt the SQL. Guard:
    if right(@sql, 1) = '+'
        select @sql = left(@sql, len(@sql) - 1);

    select @sql = @sql + ' + ''' + isnull(@DocumentFolder, '') + ''')), ''\'', ''/'') AS DocumentFolder ';

    select @sql = @sql
        + ' FROM (select c.CaseID, c.Yr, c.Orig_Atty, d.CodeVal as Case_Letter,
                                    c.Last_Name, c.First_Name, c.FileNo
                                    from vwfrmClientLedger c (nolock)
                                    inner join tblDropD d (nolock)
                                    on c.Case_Letter = d.Code
                                    where d.FieldName = ''Case_Letter'') as X
                                    WHERE CaseID = ' + convert(varchar(10), @CaseID);

    exec (@sql);

    close cursorT;
    deallocate cursorT;
end;
go


-- ---------------------------------------------------------------------------
-- 8.6 spMoveDocumentFolder (G2 rewrite) — new signature:
--       @CaseID, @OldFolderPath, @NewFolderPath
--     Updates both tblCaseDocuments AND tblScans in one transaction.
--     SET XACT_ABORT ON + TRY/CATCH guarantees no partial state. Both-zero
--     rowcount THROWs (caller must roll back the Dropbox move). The legacy
--     SP's @CaseStatus parameter is gone — the new contract puts the caller
--     in charge of computing source + destination paths (which it already
--     does, via spGetDocumentFolderName + spGetClosedDocumentFolderName).
-- ---------------------------------------------------------------------------

print N'    8.6 dropping + recreating spMoveDocumentFolder (G2)';
go

-- LEGACY (pre-Phase 4c) — preserved for rollback. The legacy body uses a
-- token-walk + position-3 _CLOSED injection that hard-codes the S:\COMMON\<Atty>\_CLIENTS\<Letter>
-- path shape. Cannot be retargeted without a rewrite, which is what G2 does.
-- ============================================================================
-- CREATE   PROCEDURE spMoveDocumentFolder
--     @CaseID INT,
--     @CaseStatus VARCHAR(20) -- Closed/Open
-- AS
-- BEGIN
--     DECLARE @DocumentRootDirectory VARCHAR(500);
--     DECLARE @SourceFileName VARCHAR(500);
--     DECLARE @TargetFileName VARCHAR(500);
--     DECLARE @CaseDocumentID AS INT;
--     DECLARE @i AS INT;
--     DECLARE @word VARCHAR(100);
--
--     SELECT @DocumentRootDirectory = DocumentRootDirectory
--     FROM dbo.tblDocumentRootDirectory;
--
--     DECLARE cursorT CURSOR READ_ONLY FAST_FORWARD FOR
--     SELECT CaseDocumentID FROM dbo.tblCaseDocuments (NOLOCK) WHERE CaseID = @CaseID;
--     OPEN cursorT;
--     WHILE (1 = 1)
--     BEGIN
--         FETCH cursorT INTO @CaseDocumentID;
--         IF @@FETCH_STATUS <> 0 BREAK;
--         SELECT @SourceFileName = SUBSTRING(DocumentFileName, LEN(@DocumentRootDirectory) + 2, LEN(DocumentFileName))
--         FROM dbo.tblCaseDocuments (NOLOCK) WHERE CaseDocumentID = @CaseDocumentID;
--
--         DECLARE cursorW CURSOR READ_ONLY FAST_FORWARD FOR
--         SELECT Word FROM dbo.fnGetListOfWords(@SourceFileName, '\') WHERE Word <> '_CLOSED';
--         SELECT @i = 1, @TargetFileName = '';
--         OPEN cursorW;
--         WHILE (1 = 1)
--         BEGIN
--             FETCH cursorW INTO @word;
--             IF @@FETCH_STATUS <> 0 BREAK;
--             IF @CaseStatus = 'Closed' AND @i = 3
--             BEGIN
--                 SELECT @TargetFileName = @TargetFileName + '\_CLOSED';
--             END;
--             SELECT @TargetFileName = @TargetFileName + '\' + @word;
--             SELECT @i = @i + 1;
--         END;
--         CLOSE cursorW;
--         DEALLOCATE cursorW;
--
--         UPDATE dbo.tblCaseDocuments
--         SET DocumentFileName = @DocumentRootDirectory + @TargetFileName
--         WHERE CaseDocumentID = @CaseDocumentID;
--     END;
--     CLOSE cursorT;
--     DEALLOCATE cursorT;
-- END;
-- ============================================================================

drop procedure if exists dbo.spMoveDocumentFolder;
go

create procedure dbo.spMoveDocumentFolder
    @CaseID         int,
    @OldFolderPath  nvarchar(500),    -- e.g., /Company/COMMON/PM/_CLIENTS/Criminal/Quail, Martha 26-139-PM/
    @NewFolderPath  nvarchar(500)     -- e.g., /Company/COMMON/PM/_CLIENTS/Criminal/_CLOSED/Quail, Martha 26-139-PM/
as
begin
    set nocount on;
    set xact_abort on;   -- any runtime error aborts the whole batch and rolls back

    -- Normalize: enforce trailing slash so prefix match is unambiguous.
    if right(@OldFolderPath, 1) <> '/' set @OldFolderPath = @OldFolderPath + '/';
    if right(@NewFolderPath, 1) <> '/' set @NewFolderPath = @NewFolderPath + '/';

    declare @CaseDocsUpdated int = 0, @ScansUpdated int = 0;

    begin try
        begin tran;

        update dbo.tblCaseDocuments
        set DocumentFileName =
            @NewFolderPath + substring(DocumentFileName, len(@OldFolderPath) + 1, len(DocumentFileName))
        where CaseID = @CaseID
          and left(DocumentFileName, len(@OldFolderPath)) = @OldFolderPath;
        set @CaseDocsUpdated = @@rowcount;

        update dbo.tblScans
        set ScanLocation =
            @NewFolderPath + substring(ScanLocation, len(@OldFolderPath) + 1, len(ScanLocation))
        where CaseID = @CaseID
          and left(ScanLocation, len(@OldFolderPath)) = @OldFolderPath;
        set @ScansUpdated = @@rowcount;

        -- Hard-fail when BOTH tables had zero matches. This means the Dropbox
        -- folder was moved but SQL has no record of the case at all, which is
        -- almost always a sign of a path mismatch or a wrong @OldFolderPath
        -- — we must not silently accept it. (One-of-two is acceptable:
        -- many cases legitimately have entries in only one of the two tables.)
        if @CaseDocsUpdated = 0 and @ScansUpdated = 0
        begin
            rollback tran;
            ;throw 51000, N'spMoveDocumentFolder: zero rows updated in both tblCaseDocuments and tblScans for the given @CaseID and @OldFolderPath. Dropbox folder was moved but SQL has no record of this case under that path — operation aborted; caller must roll back the Dropbox move.', 1;
        end

        commit tran;
    end try
    begin catch
        if xact_state() <> 0 rollback tran;
        ;throw;   -- re-raise with original error number, message, severity
    end catch;

    -- Caller (MoveDocumentByCaseStatus in DocumentManagement.bas) reads this
    -- recordset to confirm the SQL ledger matched the Dropbox move.
    -- Exactly-one-zero rowcount is logged to tblDropboxAuditLog as a warning
    -- and accepted (the case is still considered closed/reopened — common
    -- when a case has documents but no scans, or vice versa). Both-zero is
    -- impossible here because it would have raised above; the caller can
    -- rely on at least one non-zero rowcount.
    select @CaseDocsUpdated as CaseDocumentsUpdated,
           @ScansUpdated    as ScansUpdated;
end;
go


-- ---------------------------------------------------------------------------
-- 8.7 spSaveCaseDocument (G13 rewrite) — adds an unresolved-template-token
--     guard. THROWs (does NOT insert) if @DocumentName contains any of:
--       - a `[...]` substring (e.g., the [Case_Letter] defect class)
--       - the literal string `<currentdate>` (unsubstituted token)
--       - the literal string `(customuserentry)` (unsubstituted token)
--     This catches the 9-row Phase 1b defect class at its source. Legal
--     staff retain the ability to add intentional context to filenames
--     (parens for notes, dashes, etc.) — only the three template-token
--     shapes are blocked.
-- ---------------------------------------------------------------------------

print N'    8.7 dropping + recreating spSaveCaseDocument (G13)';
go

-- LEGACY (pre-Phase 4c) — preserved for rollback.
-- ============================================================================
-- CREATE   PROCEDURE [dbo].[spSaveCaseDocument]
--     @CaseID INT,
--     @DocumentType VARCHAR(250),
--     @DocumentName VARCHAR(500)
-- AS
-- BEGIN
--     -- need to delete the record if the user is using the document name for the same case
--     DELETE FROM dbo.tblCaseDocuments
--     WHERE CaseID = @CaseID
--     AND DocumentFileName = @DocumentName;
--
--     INSERT INTO dbo.tblCaseDocuments (CaseID, DocumentType, DocumentFileName)
--     SELECT @CaseID, @DocumentType, @DocumentName;
-- END;
-- ============================================================================

drop procedure if exists dbo.spSaveCaseDocument;
go

create procedure dbo.spSaveCaseDocument
    @CaseID int,
    @DocumentType varchar(250),
    @DocumentName varchar(500)
as
begin
    set nocount on;

    -- G13 guard: reject unresolved template tokens. Catches the 9-row
    -- Phase 1b defect class (e.g., [Case_Letter] surviving into the stored
    -- path because the vwfrmClientLedger row was incomplete at save time).
    -- The pattern '%[[]%]%' matches any string with a [ followed somewhere
    -- by a ] — covers any `[Token]` shape regardless of token name.
    if @DocumentName like '%[[]%]%'
       or charindex('<currentdate>', @DocumentName) > 0
       or charindex('(customuserentry)', @DocumentName) > 0
    begin
        ;throw 51001, N'spSaveCaseDocument: @DocumentName contains unresolved template tokens (e.g., [Field], <currentdate>, or (customuserentry)). Resolve the tokens before saving — see plan G13.', 1;
    end;

    -- Dedupe: the user may re-save the same DocumentName for the same case
    -- (e.g., re-scan, overwrite); delete any existing row before inserting.
    delete from dbo.tblCaseDocuments
    where CaseID = @CaseID
      and DocumentFileName = @DocumentName;

    insert into dbo.tblCaseDocuments
    (
        CaseID,
        DocumentType,
        DocumentFileName
    )
    select @CaseID,
           @DocumentType,
           @DocumentName;
end;
go


-- ---------------------------------------------------------------------------
-- 8.8 Verification — exec each path-building SP for the known-good
--     QUAIL/30337 case and print the resolved /Company/ path. After install,
--     the IT admin can eyeball the output to confirm forward-slash form.
-- ---------------------------------------------------------------------------

print N'';
print N'>>> SECTION 8 verification — exec path-building SPs for QUAIL/30337';
go

print N'';
print N'    spGetIntakeFolderName():';
exec dbo.spGetIntakeFolderName;
go

print N'';
print N'    spGetDocumentFolderName(''General'', 30337):';
exec dbo.spGetDocumentFolderName 'General', 30337;
go

print N'';
print N'    spGetClosedDocumentFolderName(''General'', 30337):';
exec dbo.spGetClosedDocumentFolderName 'General', 30337;
go

print N'';
print N'    spGetClosedFileScanFolderName(''General'', 30337):';
exec dbo.spGetClosedFileScanFolderName 'General', 30337;
go

print N'';
print N'    spGetAllInvoicesFolderName(30337):';
exec dbo.spGetAllInvoicesFolderName 30337;
go

print N'    SECTION 8 done.';
go


-- #############################################################################
-- SECTION 9 — DROPBOX BRIDGE SERVICE OBJECTS (Phase A of the bridge plan)
-- #############################################################################
--   Adds the two SQL objects the TBCMSDropboxBridge service depends on:
--     9.1  tblDropboxConfig.BridgeUrl  — VBA reads the bridge URL from here at
--          startup so it can change without recompiling the .accde.
--     9.2  tblDropboxServiceToken      — singleton row holding the one service-
--          account OAuth token (Data-Protection-encrypted by the bridge). VBA
--          never touches this table; the bridge reads and UPSERTs it.
--   Both are idempotent (IF NOT EXISTS guarded).
--
--   RE-RUN CAVEAT: SECTION 1.2 DROP/CREATEs tblDropboxConfig, so a destructive
--   re-run drops the BridgeUrl column; 9.1 (running later in the same pass)
--   re-adds and re-seeds it to the placeholder URL — IT must re-apply the real
--   URL afterward, exactly as for AppSecret.
-- #############################################################################

print N'';
print N'>>> SECTION 9: Dropbox Bridge service objects';
go

-- ---------------------------------------------------------------------------
-- 9.1 — BridgeUrl column on tblDropboxConfig
-- ---------------------------------------------------------------------------
if not exists
(
    select 1
    from sys.columns
    where object_id = object_id('dbo.tblDropboxConfig')
          and name = 'BridgeUrl'
)
begin
    alter table dbo.tblDropboxConfig
        add BridgeUrl nvarchar(500) null;

    -- separate batch needed before the column is referenceable in this script
    exec ('update dbo.tblDropboxConfig
           set    BridgeUrl = N''http://tbcms-bridge.tatebywater.local/api''
           where  ConfigID = 1;');

    print N'    SECTION 9.1: BridgeUrl column added and seeded (placeholder URL).';
end;
else
    print N'    SECTION 9.1: BridgeUrl already present — skipped.';
go

-- ---------------------------------------------------------------------------
-- 9.2 — tblDropboxServiceToken (singleton service-account token row)
-- ---------------------------------------------------------------------------
if not exists
(
    select 1
    from sys.tables
    where object_id = object_id('dbo.tblDropboxServiceToken')
)
begin
    create table dbo.tblDropboxServiceToken
    (
        TokenID int not null primary key,                 -- always 1 (singleton)
        AccessToken nvarchar(max) not null,               -- Data-Protection-encrypted (machine scope)
        RefreshToken nvarchar(max) not null,              -- Data-Protection-encrypted (machine scope)
        ExpiresAtUtc datetime2 not null,
        AccountEmail nvarchar(200) null,
        UpdatedAtUtc datetime2 not null
            constraint DF_DropboxServiceToken_UpdatedAtUtc
                default (sysutcdatetime()),
        SetupByUser nvarchar(200) null,                   -- Windows login that ran setup
        -- Single-row guarantee — mirrors tblDropboxConfig's CK_..._SingleRow.
        -- The bridge UPSERTs TokenID = 1; it must never accumulate rows.
        constraint CK_DropboxServiceToken_SingleRow check (TokenID = 1)
    );

    print N'    SECTION 9.2: tblDropboxServiceToken created.';
end;
else
    print N'    SECTION 9.2: tblDropboxServiceToken already present — skipped.';
go

print N'    SECTION 9 done.';
go


print N'';
print N'================================================================================';
print N'INSTALL_all.sql — complete.';
print N'  Don''t forget: UPDATE dbo.tblDropboxConfig SET AppSecret = N''<real>'' WHERE ConfigID = 1;';
print N'  Bridge: UPDATE dbo.tblDropboxConfig SET BridgeUrl = N''http://<server>/api'' WHERE ConfigID = 1;';
print N'  Review SECTION 7 output for B-4, B-6, B-7, B-8, B-9 triage decisions.';
print N'================================================================================';
go
