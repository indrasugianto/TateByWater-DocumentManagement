-- =============================================================================
-- STEP_1_UPDATE.sql
--
-- PURPOSE  : Migrate all DocumentFileName paths in tblCaseDocuments and
--            ScanLocation paths in tblScans from S:\ UNC format to Dropbox
--            /Company/ path format.
--
-- BEFORE RUNNING:
--   1. Run STEP_0_ANALYZE.sql and confirm all output looks expected.
--   2. Take a database backup or note the current transaction log position.
--   3. Verify STEP_0_ANALYZE query 4 returns zero rows (no unexpected prefixes).
--   4. Verify STEP_0_ANALYZE query 3 rows — decide whether to update template
--      literal rows or leave them (they will be updated by this script; see note).
--
-- WHAT THIS SCRIPT DOES:
--   For every non-null, non-empty DocumentFileName / ScanLocation value:
--     a) Strip a leading '#' if present (Access hyperlink format artifact)
--     b) Replace 'S:\' with '/Company/'
--     c) Replace all remaining '\' with '/'
--
--   Example transformations:
--     S:\COMMON\RLF\_CLIENTS\Domestic\Smith D23-001\file.pdf
--     → /Company/COMMON/RLF/_CLIENTS/Domestic/Smith D23-001/file.pdf
--
--     S:\Closed File Scans\Smith D23-001\scan.pdf
--     → /Company/Closed File Scans/Smith D23-001/scan.pdf
--
--     #S:\COMMON\RLF\_CLIENTS\file.pdf  (hyperlink artifact)
--     → /Company/COMMON/RLF/_CLIENTS/file.pdf
--
--   NOTE on template literal rows (e.g. S:\COMMON\RLF\CLIENTS\[Case_Letter]\...):
--   These 7 rows are already broken and cannot be resolved by this migration.
--   They are updated using the same formula — the [Field] tokens remain in the
--   path but the root is corrected. They must be reviewed and fixed manually
--   after cutover. Their CaseDocumentIDs are listed in STEP_0_ANALYZE query 3.
--
-- SERVER   : awsql2022dev
-- DATABASE : TateByWater
-- =============================================================================

USE TateByWater;
GO

BEGIN TRANSACTION MigratePathsToDropbox;

-- ---------------------------------------------------------------------------
-- PART A: tblCaseDocuments.DocumentFileName
-- ---------------------------------------------------------------------------

DECLARE @CountDocsBefore  INT;
DECLARE @CountDocsUpdated INT;

SELECT @CountDocsBefore = COUNT(*)
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> '';

UPDATE tblCaseDocuments
SET DocumentFileName =
    REPLACE(
        REPLACE(
            -- Strip leading '#' (Access hyperlink format)
            CASE WHEN LEFT(DocumentFileName, 1) = '#'
                 THEN SUBSTRING(DocumentFileName, 2, LEN(DocumentFileName))
                 ELSE DocumentFileName
            END,
            'S:\', '/Company/'   -- root replacement
        ),
        '\', '/'                 -- normalize remaining backslashes
    )
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
  AND (
        DocumentFileName LIKE 'S:\%'
     OR DocumentFileName LIKE '#S:\%'
  );

SET @CountDocsUpdated = @@ROWCOUNT;

-- ---------------------------------------------------------------------------
-- PART B: tblScans.ScanLocation
-- ---------------------------------------------------------------------------

DECLARE @CountScansBefore  INT;
DECLARE @CountScansUpdated INT;

SELECT @CountScansBefore = COUNT(*)
FROM tblScans
WHERE ScanLocation IS NOT NULL
  AND ScanLocation <> '';

UPDATE tblScans
SET ScanLocation =
    REPLACE(
        REPLACE(
            CASE WHEN LEFT(ScanLocation, 1) = '#'
                 THEN SUBSTRING(ScanLocation, 2, LEN(ScanLocation))
                 ELSE ScanLocation
            END,
            'S:\', '/Company/'
        ),
        '\', '/'
    )
WHERE ScanLocation IS NOT NULL
  AND ScanLocation <> ''
  AND (
        ScanLocation LIKE 'S:\%'
     OR ScanLocation LIKE '#S:\%'
  );

SET @CountScansUpdated = @@ROWCOUNT;

-- ---------------------------------------------------------------------------
-- SUMMARY
-- ---------------------------------------------------------------------------

SELECT
    'tblCaseDocuments' AS TableName,
    @CountDocsBefore   AS RowsWithPaths,
    @CountDocsUpdated  AS RowsUpdated;

SELECT
    'tblScans'         AS TableName,
    @CountScansBefore  AS RowsWithPaths,
    @CountScansUpdated AS RowsUpdated;

-- ---------------------------------------------------------------------------
-- SPOT-CHECK: sample of updated rows — verify they look correct before COMMIT
-- ---------------------------------------------------------------------------

SELECT TOP 20
    CaseDocumentID,
    DocumentFileName AS UpdatedPath
FROM tblCaseDocuments
WHERE DocumentFileName LIKE '/Company/%'
ORDER BY CaseDocumentID;

SELECT TOP 10
    ScanID,
    ScanLocation AS UpdatedPath
FROM tblScans
WHERE ScanLocation LIKE '/Company/%'
ORDER BY ScanID;

-- ---------------------------------------------------------------------------
-- DECISION POINT
--   If the spot-check rows above look correct → run:  COMMIT TRANSACTION MigratePathsToDropbox;
--   If anything looks wrong              → run:  ROLLBACK TRANSACTION MigratePathsToDropbox;
--
--   DO NOT run both. Pick one.
-- ---------------------------------------------------------------------------

-- COMMIT TRANSACTION MigratePathsToDropbox;
-- ROLLBACK TRANSACTION MigratePathsToDropbox;
