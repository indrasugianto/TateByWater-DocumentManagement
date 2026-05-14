-- =============================================================================
-- STEP_1_UPDATE.sql
--
-- PURPOSE  : Rewrite legacy S:\-rooted (and equivalent UNC / Access-hyperlink /
--            URL-encoded) paths to Dropbox /Company/ paths across:
--              (A) tblCaseDocuments.DocumentFileName
--              (B) tblScans.ScanLocation
--              (C) [TB Intakes].[Scan Location GI]
--
--            Rows that cannot be rewritten safely are copied into per-table
--            manual-triage tables for hand fixing. The script is idempotent:
--            re-running it is a no-op for already-migrated rows.
--
-- BEFORE RUNNING:
--   1. Run STEP_0_ANALYZE.sql and review its output.
--   2. Take a database backup (or note the current LSN).
--   3. Confirm the matched-categories sum equals the populated row count for
--      each table — see the SUMMARY block at the bottom of this script.
--
-- WHAT THIS SCRIPT DOES (per-category transforms applied in this order):
--
--   For each path value (X) the script first removes the Access hyperlink
--   wrapper, then collapses every legacy root to /Company/, then normalizes
--   backslashes. Categories handled:
--
--     A-1  X = 'S:\…'                       (tblCaseDocuments, all 26,043 rows)
--     B-1  X = '#S:\…#'                     (Access hyperlink, 3,879 rows)
--     B-2  X = 'S:\…'                       (bare path, 26 rows)
--     B-3  X = '#file:///S:\…#'             (URL-encoded with file:// wrapper, 617 rows)
--     B-4  X = '#?S:\…#'                    (typo '?' prefix, 60 rows)
--     B-5  X = '#file:///\\TBF-SRVR12\<co>\…#'  (legacy UNC + file:// wrapper, 19 rows)
--     B-6  X = '#\\TBF-SRVR12\<co>\…#'      (legacy UNC, bare, 6 rows)
--     C-1  X = 'S:\…'                       ([TB Intakes], 849 rows)
--     C-2  X = '#S:\…#'                     (43 rows)
--     C-3  X = '#file:///S:\…#'             (57 rows)
--     C-4  X = '?S:\…'                      (1 row)
--
--   Transformations applied per category:
--     - Strip leading '#' and trailing '#' (Access hyperlink wrapping)
--     - Strip leading 'file:///' if present
--     - Strip leading '?' (typo) if present
--     - Replace '\\TBF-SRVR12\Company\' with '/Company/' (legacy UNC ≡ S:)
--       Note: DB collation SQL_Latin1_General_CP1_CI_AS makes REPLACE case
--       insensitive, so both 'company' (lowercase) and 'Company' (mixed) match.
--     - URL-decode (only '%20' appears in the data — verified 2026-05-14)
--     - Replace 'S:\' with '/Company/'
--     - Replace remaining '\' with '/'
--
-- ROWS NOT REWRITTEN (quarantined for manual triage):
--
--   - tblScans: 9 'OTHER' rows that are hash-less, mid-string corrupted, or
--     missing the 'S:' root entirely. Copied to dbo.tblScans_ManualTriage.
--   - [TB Intakes]: 1 'OTHER' row missing the 'S:' root. Copied to
--     dbo.TBIntakes_ManualTriage.
--   - tblCaseDocuments: 13 rows where DocumentFileName contains a literal '['.
--     These split into two defects (see Phase 1b in the migration plan):
--       9 rows with unresolved template tokens (e.g. '[Case_Letter]') —
--         require re-running spGetDocumentFolderName + spGetDocumentFileName
--         after the vwfrmClientLedger row is repopulated.
--       4 rows ending in '.[df' (truncated '.pdf' with a stray '['); need
--         the on-disk file inspected to recover the correct filename.
--     This script rewrites their root (S:\ → /Company/) but leaves the
--     defective tokens / extensions intact. They are listed by ID in
--     STEP_0_ANALYZE query 3 and STEP_2_VERIFY query 7.
--
-- INTAKE PATH MAPPING (note for reviewers):
--   This script does a like-for-like rewrite of intake paths. Existing
--   intakes live at S:\Closed File Scans\TB\Intakes\… and will rewrite to
--   /Company/Closed File Scans/TB/Intakes/…. The migration plan's
--   tblDropboxRootConfig.IntakeDirectory proposes /Company/COMMON/Intakes —
--   that is a structural relocation, not a path rewrite, and is out of scope
--   for this script. Relocation (if approved) is a separate one-off SQL
--   that runs AFTER this script.
--
-- SERVER   : awsql2022dev (test environment for Phases 0–6)
-- DATABASE : TateByWater
-- =============================================================================

USE TateByWater;
GO

-- ---------------------------------------------------------------------------
-- 0. Manual-triage tables (create if missing — preserved across re-runs)
-- ---------------------------------------------------------------------------

IF OBJECT_ID('dbo.tblScans_ManualTriage', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.tblScans_ManualTriage (
        ScansID         INT          NOT NULL PRIMARY KEY,
        OriginalValue   VARCHAR(MAX) NOT NULL,
        Reason          NVARCHAR(200) NOT NULL,
        QuarantinedAt   DATETIME     NOT NULL DEFAULT SYSDATETIME()
    );
END;

IF OBJECT_ID('dbo.TBIntakes_ManualTriage', 'U') IS NULL
BEGIN
    -- TB Intakes has no surrogate PK; use the natural-key columns and the value.
    CREATE TABLE dbo.TBIntakes_ManualTriage (
        TriageID        INT IDENTITY(1,1) PRIMARY KEY,
        GILastName      VARCHAR(255) NULL,
        GIFirstName     VARCHAR(255) NULL,
        GIDate          DATETIME     NULL,
        OriginalValue   VARCHAR(MAX) NOT NULL,
        Reason          NVARCHAR(200) NOT NULL,
        QuarantinedAt   DATETIME     NOT NULL DEFAULT SYSDATETIME()
    );
END;
GO

BEGIN TRANSACTION MigratePathsToDropbox;

-- =============================================================================
-- PART A — tblCaseDocuments.DocumentFileName
-- =============================================================================
-- All 26,043 rows start with 'S:\' (verified 2026-05-14). The leading-'#' and
-- trailing-'#' branches are kept as defense in depth in case production drifts.

DECLARE @CountDocsBefore  INT;
DECLARE @CountDocsUpdated INT;

SELECT @CountDocsBefore = COUNT(*)
FROM dbo.tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> '';

UPDATE dbo.tblCaseDocuments
SET DocumentFileName =
    REPLACE(
        REPLACE(
            -- Strip leading '#' and trailing '#' if present.
            CASE
                WHEN LEFT(DocumentFileName,1)='#' AND RIGHT(DocumentFileName,1)='#'
                    THEN SUBSTRING(DocumentFileName, 2, LEN(DocumentFileName) - 2)
                WHEN LEFT(DocumentFileName,1)='#'
                    THEN SUBSTRING(DocumentFileName, 2, LEN(DocumentFileName) - 1)
                WHEN RIGHT(DocumentFileName,1)='#'
                    THEN LEFT(DocumentFileName, LEN(DocumentFileName) - 1)
                ELSE DocumentFileName
            END,
            'S:\', '/Company/'),
        '\', '/')
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
  AND (DocumentFileName LIKE 'S:\%' OR DocumentFileName LIKE '#S:\%');

SET @CountDocsUpdated = @@ROWCOUNT;

-- =============================================================================
-- PART B — tblScans.ScanLocation (multi-pass per category)
-- =============================================================================

DECLARE @CountScansBefore   INT;
DECLARE @CountScansB1       INT = 0;
DECLARE @CountScansB2       INT = 0;
DECLARE @CountScansB3       INT = 0;
DECLARE @CountScansB4       INT = 0;
DECLARE @CountScansB5       INT = 0;
DECLARE @CountScansB6       INT = 0;
DECLARE @CountScansQuarantined INT = 0;

SELECT @CountScansBefore = COUNT(*)
FROM dbo.tblScans
WHERE ScanLocation IS NOT NULL AND ScanLocation <> '';

-- B-1: '#S:\…#'  — Access hyperlink wrapper around a bare S:\ path.
--      A few rows (2 of 3,880) have '%20' encoded inside the path even though
--      they lack the 'file:///' wrapper; URL-decode is applied unconditionally.
UPDATE dbo.tblScans
SET ScanLocation =
    REPLACE(
        REPLACE(
            REPLACE(
                -- strip leading '#' and trailing '#'
                SUBSTRING(ScanLocation, 2, LEN(ScanLocation) -
                    CASE WHEN RIGHT(ScanLocation,1)='#' THEN 2 ELSE 1 END),
                '%20', ' '),
            'S:\', '/Company/'),
        '\', '/')
WHERE ScanLocation LIKE '#S:\%';
SET @CountScansB1 = @@ROWCOUNT;

-- B-2: 'S:\…path#\\TBF-SRVR12\…#'  — full Access hyperlink whose displaytext
--      starts with 'S:\' and is followed by a UNC-formatted URL part. The
--      simple LIKE 'S:\%' filter would have suggested these are bare paths,
--      but every one of the 26 rows actually contains a '#…#' suffix.
--      Truncate at the first '#' to keep the displaytext (the S:\ path).
UPDATE dbo.tblScans
SET ScanLocation =
    REPLACE(
        REPLACE(
            REPLACE(
                CASE WHEN CHARINDEX('#', ScanLocation) > 0
                     THEN LEFT(ScanLocation, CHARINDEX('#', ScanLocation) - 1)
                     ELSE ScanLocation END,
                '%20', ' '),
            'S:\', '/Company/'),
        '\', '/')
WHERE ScanLocation LIKE 'S:\%';
SET @CountScansB2 = @@ROWCOUNT;

-- B-3: '#file:///S:\…#' — URL-encoded with file:// wrapper.
--      Strip '#…#' wrap → strip 'file:///' → URL-decode → S:\ → /Company/ → \ → /
UPDATE dbo.tblScans
SET ScanLocation =
    REPLACE(
        REPLACE(
            REPLACE(
                -- After stripping '#…#' the value starts 'file:///S:\…',
                -- which is 8 chars of 'file:///' prefix to remove.
                SUBSTRING(ScanLocation, 2 + 8, LEN(ScanLocation) -
                    CASE WHEN RIGHT(ScanLocation,1)='#' THEN 2 ELSE 1 END - 8),
                '%20', ' '),
            'S:\', '/Company/'),
        '\', '/')
WHERE ScanLocation LIKE '#file:///S:\%';
SET @CountScansB3 = @@ROWCOUNT;

-- B-4: '#?S:\…#' — '#?' typo prefix.
UPDATE dbo.tblScans
SET ScanLocation =
    REPLACE(
        REPLACE(
            -- Strip leading '#?' (2 chars) and trailing '#' (1 char).
            SUBSTRING(ScanLocation, 3, LEN(ScanLocation) -
                CASE WHEN RIGHT(ScanLocation,1)='#' THEN 3 ELSE 2 END),
            'S:\', '/Company/'),
        '\', '/')
WHERE ScanLocation LIKE '#?S:\%';
SET @CountScansB4 = @@ROWCOUNT;

-- B-5: '#file:///\\TBF-SRVR12\<co>\…#' — legacy UNC + file:// wrapper.
--      Strip '#…#' wrap → strip 'file:///' → replace UNC root with /Company/
--      → URL-decode → \ → /
UPDATE dbo.tblScans
SET ScanLocation =
    REPLACE(
        REPLACE(
            REPLACE(
                REPLACE(
                    SUBSTRING(ScanLocation, 2 + 8, LEN(ScanLocation) -
                        CASE WHEN RIGHT(ScanLocation,1)='#' THEN 2 ELSE 1 END - 8),
                    '%20', ' '),
                '\\TBF-SRVR12\Company\', '/Company/'),
            'S:\', '/Company/'),
        '\', '/')
WHERE ScanLocation LIKE '#file:///\\TBF-SRVR12\%';
SET @CountScansB5 = @@ROWCOUNT;

-- B-6: '#\\TBF-SRVR12\<co>\…#' — legacy UNC, bare.
UPDATE dbo.tblScans
SET ScanLocation =
    REPLACE(
        REPLACE(
            REPLACE(
                SUBSTRING(ScanLocation, 2, LEN(ScanLocation) -
                    CASE WHEN RIGHT(ScanLocation,1)='#' THEN 2 ELSE 1 END),
                '\\TBF-SRVR12\Company\', '/Company/'),
            'S:\', '/Company/'),
        '\', '/')
WHERE ScanLocation LIKE '#\\TBF-SRVR12\%';
SET @CountScansB6 = @@ROWCOUNT;

-- B-7: Quarantine — anything non-null/non-empty that matches none of B-1..B-6
--      and is not already migrated (does not start with '/Company/').
INSERT INTO dbo.tblScans_ManualTriage (ScansID, OriginalValue, Reason)
SELECT s.ScansID,
       s.ScanLocation,
       'Hash-less or mid-string-corrupted ScanLocation — no automatic rewrite'
FROM dbo.tblScans s
WHERE s.ScanLocation IS NOT NULL
  AND LTRIM(RTRIM(s.ScanLocation)) <> ''
  AND s.ScanLocation NOT LIKE '/Company/%'
  AND s.ScanLocation NOT LIKE '#S:\%'
  AND s.ScanLocation NOT LIKE 'S:\%'
  AND s.ScanLocation NOT LIKE '#file:///S:\%'
  AND s.ScanLocation NOT LIKE '#?S:\%'
  AND s.ScanLocation NOT LIKE '#file:///\\TBF-SRVR12\%'
  AND s.ScanLocation NOT LIKE '#\\TBF-SRVR12\%'
  AND NOT EXISTS (
      SELECT 1 FROM dbo.tblScans_ManualTriage q WHERE q.ScansID = s.ScansID);
SET @CountScansQuarantined = @@ROWCOUNT;

-- =============================================================================
-- PART C — [TB Intakes].[Scan Location GI] (multi-pass per category)
-- =============================================================================

DECLARE @CountIntakesBefore       INT;
DECLARE @CountIntakesC1           INT = 0;
DECLARE @CountIntakesC2           INT = 0;
DECLARE @CountIntakesC3           INT = 0;
DECLARE @CountIntakesC4           INT = 0;
DECLARE @CountIntakesQuarantined  INT = 0;

SELECT @CountIntakesBefore = COUNT(*)
FROM [TB Intakes]
WHERE [Scan Location GI] IS NOT NULL AND [Scan Location GI] <> '';

-- C-1: 'S:\…'  — bare path.
UPDATE [TB Intakes]
SET [Scan Location GI] =
    REPLACE(REPLACE([Scan Location GI], 'S:\', '/Company/'), '\', '/')
WHERE [Scan Location GI] LIKE 'S:\%';
SET @CountIntakesC1 = @@ROWCOUNT;

-- C-2: '#S:\…#'  — Access hyperlink wrapper.
UPDATE [TB Intakes]
SET [Scan Location GI] =
    REPLACE(
        REPLACE(
            SUBSTRING([Scan Location GI], 2, LEN([Scan Location GI]) -
                CASE WHEN RIGHT([Scan Location GI],1)='#' THEN 2 ELSE 1 END),
            'S:\', '/Company/'),
        '\', '/')
WHERE [Scan Location GI] LIKE '#S:\%';
SET @CountIntakesC2 = @@ROWCOUNT;

-- C-3: '#file:///S:\…#'  — URL-encoded with file:// wrapper.
UPDATE [TB Intakes]
SET [Scan Location GI] =
    REPLACE(
        REPLACE(
            REPLACE(
                SUBSTRING([Scan Location GI], 2 + 8, LEN([Scan Location GI]) -
                    CASE WHEN RIGHT([Scan Location GI],1)='#' THEN 2 ELSE 1 END - 8),
                '%20', ' '),
            'S:\', '/Company/'),
        '\', '/')
WHERE [Scan Location GI] LIKE '#file:///S:\%';
SET @CountIntakesC3 = @@ROWCOUNT;

-- C-4: '?S:\…'  — '?' typo prefix, no Access hyperlink wrapper.
UPDATE [TB Intakes]
SET [Scan Location GI] =
    REPLACE(
        REPLACE(
            SUBSTRING([Scan Location GI], 2, LEN([Scan Location GI]) - 1),
            'S:\', '/Company/'),
        '\', '/')
WHERE [Scan Location GI] LIKE '?S:\%';
SET @CountIntakesC4 = @@ROWCOUNT;

-- C-5: Quarantine — non-null, non-already-migrated rows that match no pattern.
INSERT INTO dbo.TBIntakes_ManualTriage
    (GILastName, GIFirstName, GIDate, OriginalValue, Reason)
SELECT i.[GI Last Name], i.[GI First Name], i.[GI Date],
       i.[Scan Location GI],
       'Hash-less or root-missing Scan Location GI — no automatic rewrite'
FROM [TB Intakes] i
WHERE i.[Scan Location GI] IS NOT NULL
  AND LTRIM(RTRIM(i.[Scan Location GI])) <> ''
  AND i.[Scan Location GI] NOT LIKE '/Company/%'
  AND i.[Scan Location GI] NOT LIKE 'S:\%'
  AND i.[Scan Location GI] NOT LIKE '#S:\%'
  AND i.[Scan Location GI] NOT LIKE '#file:///S:\%'
  AND i.[Scan Location GI] NOT LIKE '?S:\%'
  AND NOT EXISTS (
      SELECT 1 FROM dbo.TBIntakes_ManualTriage q
       WHERE q.OriginalValue = i.[Scan Location GI]
         AND ISNULL(q.GILastName,'')  = ISNULL(i.[GI Last Name],'')
         AND ISNULL(q.GIFirstName,'') = ISNULL(i.[GI First Name],''));
SET @CountIntakesQuarantined = @@ROWCOUNT;

-- =============================================================================
-- SUMMARY — counts per pass + leftover-offender sanity check
-- =============================================================================

SELECT 'tblCaseDocuments' AS TableName,
       @CountDocsBefore   AS RowsWithPaths,
       @CountDocsUpdated  AS RowsUpdated;

SELECT 'tblScans'                              AS TableName,
       @CountScansBefore                       AS RowsWithPaths,
       @CountScansB1                           AS Pass_B1_HashSPath,
       @CountScansB2                           AS Pass_B2_BareSPath,
       @CountScansB3                           AS Pass_B3_FileURL_SPath,
       @CountScansB4                           AS Pass_B4_HashQuestionSPath,
       @CountScansB5                           AS Pass_B5_FileURL_UNC,
       @CountScansB6                           AS Pass_B6_HashUNC,
       @CountScansQuarantined                  AS Quarantined,
       @CountScansB1 + @CountScansB2 + @CountScansB3
       + @CountScansB4 + @CountScansB5 + @CountScansB6
       + @CountScansQuarantined                AS Accounted;

SELECT 'TB Intakes'                            AS TableName,
       @CountIntakesBefore                     AS RowsWithPaths,
       @CountIntakesC1                         AS Pass_C1_BareSPath,
       @CountIntakesC2                         AS Pass_C2_HashSPath,
       @CountIntakesC3                         AS Pass_C3_FileURL_SPath,
       @CountIntakesC4                         AS Pass_C4_QuestionSPath,
       @CountIntakesQuarantined                AS Quarantined,
       @CountIntakesC1 + @CountIntakesC2 + @CountIntakesC3
       + @CountIntakesC4 + @CountIntakesQuarantined AS Accounted;

-- Leftover-offender check. Expected: every row = 0.
-- Quarantined rows are excluded so they don't count as "leftovers" — they're
-- intentionally untouched and listed in their respective triage tables.
SELECT 'tblCaseDocuments' AS TableName,
       SUM(CASE WHEN DocumentFileName LIKE '%\%'   THEN 1 ELSE 0 END) AS HasBackslash,
       SUM(CASE WHEN DocumentFileName LIKE '%S:%'  THEN 1 ELSE 0 END) AS HasSColon,
       SUM(CASE WHEN LEFT(DocumentFileName,1) = '#' THEN 1 ELSE 0 END) AS LeadingHash,
       SUM(CASE WHEN RIGHT(DocumentFileName,1)= '#' THEN 1 ELSE 0 END) AS TrailingHash
FROM dbo.tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> '';

SELECT 'tblScans' AS TableName,
       SUM(CASE WHEN s.ScanLocation LIKE '%\%'    THEN 1 ELSE 0 END) AS HasBackslash,
       SUM(CASE WHEN s.ScanLocation LIKE '%S:%'   THEN 1 ELSE 0 END) AS HasSColon,
       SUM(CASE WHEN LEFT(s.ScanLocation,1) = '#'  THEN 1 ELSE 0 END) AS LeadingHash,
       SUM(CASE WHEN RIGHT(s.ScanLocation,1)= '#'  THEN 1 ELSE 0 END) AS TrailingHash
FROM dbo.tblScans s
WHERE s.ScanLocation IS NOT NULL
  AND s.ScanLocation <> ''
  AND NOT EXISTS (SELECT 1 FROM dbo.tblScans_ManualTriage q WHERE q.ScansID = s.ScansID);

SELECT 'TB Intakes' AS TableName,
       SUM(CASE WHEN [Scan Location GI] LIKE '%\%'   THEN 1 ELSE 0 END) AS HasBackslash,
       SUM(CASE WHEN [Scan Location GI] LIKE '%S:%'  THEN 1 ELSE 0 END) AS HasSColon,
       SUM(CASE WHEN LEFT([Scan Location GI],1) = '#' THEN 1 ELSE 0 END) AS LeadingHash,
       SUM(CASE WHEN RIGHT([Scan Location GI],1)= '#' THEN 1 ELSE 0 END) AS TrailingHash
FROM [TB Intakes]
WHERE [Scan Location GI] IS NOT NULL
  AND [Scan Location GI] <> ''
  AND NOT EXISTS (
      SELECT 1 FROM dbo.TBIntakes_ManualTriage q
       WHERE q.OriginalValue = [TB Intakes].[Scan Location GI]);

-- ---------------------------------------------------------------------------
-- SPOT-CHECK: sample of updated rows — verify they look correct before COMMIT
-- ---------------------------------------------------------------------------

SELECT TOP 20
    CaseDocumentID,
    DocumentFileName AS UpdatedPath
FROM dbo.tblCaseDocuments
WHERE DocumentFileName LIKE '/Company/%'
ORDER BY CaseDocumentID;

SELECT TOP 10
    ScansID,
    ScanLocation AS UpdatedPath
FROM dbo.tblScans
WHERE ScanLocation LIKE '/Company/%'
ORDER BY ScansID;

SELECT TOP 10
    [GI Last Name],
    [GI First Name],
    [Scan Location GI] AS UpdatedPath
FROM [TB Intakes]
WHERE [Scan Location GI] LIKE '/Company/%'
ORDER BY [GI Last Name];

-- Manual-triage reports (review by hand; rewrite or delete per case):
SELECT * FROM dbo.tblScans_ManualTriage      ORDER BY ScansID;
SELECT * FROM dbo.TBIntakes_ManualTriage     ORDER BY TriageID;

-- ---------------------------------------------------------------------------
-- DECISION POINT
--   If every "Accounted" equals its "RowsWithPaths", every leftover-offender
--   row reads 0/0/0/0, and the spot-check rows look correct →
--       COMMIT TRANSACTION MigratePathsToDropbox;
--
--   Otherwise →
--       ROLLBACK TRANSACTION MigratePathsToDropbox;
--
--   DO NOT run both. Pick one.
-- ---------------------------------------------------------------------------

-- COMMIT   TRANSACTION MigratePathsToDropbox;
-- ROLLBACK TRANSACTION MigratePathsToDropbox;
