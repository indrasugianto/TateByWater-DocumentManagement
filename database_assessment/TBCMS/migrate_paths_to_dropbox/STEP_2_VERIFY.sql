-- =============================================================================
-- STEP_2_VERIFY.sql
--
-- PURPOSE  : Post-migration verification. Run after STEP_1_UPDATE is committed.
--            Expected: zero rows in the "Remaining S:\ paths" queries.
--
-- SERVER   : awsql2022dev
-- DATABASE : TateByWater
-- =============================================================================

USE TateByWater;
GO

-- ---------------------------------------------------------------------------
-- 1. Confirm no S:\ or #S:\ paths remain in tblCaseDocuments
--    Expected: 0 rows
-- ---------------------------------------------------------------------------
SELECT
    CaseDocumentID,
    DocumentFileName
FROM tblCaseDocuments
WHERE DocumentFileName LIKE 'S:\%'
   OR DocumentFileName LIKE '#S:\%';
GO

-- ---------------------------------------------------------------------------
-- 2. Confirm no S:\ or #S:\ paths remain in tblScans
--    Expected: 0 rows
-- ---------------------------------------------------------------------------
SELECT
    ScanID,
    ScanLocation
FROM tblScans
WHERE ScanLocation LIKE 'S:\%'
   OR ScanLocation LIKE '#S:\%';
GO

-- ---------------------------------------------------------------------------
-- 3. tblCaseDocuments path prefix distribution (should be all /Company/...)
-- ---------------------------------------------------------------------------
SELECT
    LEFT(DocumentFileName, 30)  AS PathPrefix,
    COUNT(*)                    AS RowCount
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
GROUP BY LEFT(DocumentFileName, 30)
ORDER BY COUNT(*) DESC;
GO

-- ---------------------------------------------------------------------------
-- 4. tblScans path prefix distribution
-- ---------------------------------------------------------------------------
SELECT
    LEFT(ScanLocation, 30)  AS PathPrefix,
    COUNT(*)                AS RowCount
FROM tblScans
WHERE ScanLocation IS NOT NULL
  AND ScanLocation <> ''
GROUP BY LEFT(ScanLocation, 30)
ORDER BY COUNT(*) DESC;
GO

-- ---------------------------------------------------------------------------
-- 5. Rows that still contain backslashes (should be 0 after migration)
-- ---------------------------------------------------------------------------
SELECT
    'tblCaseDocuments'  AS TableName,
    COUNT(*)            AS RowsWithBackslash
FROM tblCaseDocuments
WHERE DocumentFileName LIKE '%\%'

UNION ALL

SELECT
    'tblScans'          AS TableName,
    COUNT(*)            AS RowsWithBackslash
FROM tblScans
WHERE ScanLocation LIKE '%\%';
GO

-- ---------------------------------------------------------------------------
-- 6. Sample of migrated paths for human review (20 rows each table)
-- ---------------------------------------------------------------------------
SELECT TOP 20
    CaseDocumentID,
    CaseID,
    DocumentFileName
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
ORDER BY CaseDocumentID;
GO

SELECT TOP 20
    ScanID,
    ScanLocation
FROM tblScans
WHERE ScanLocation IS NOT NULL
  AND ScanLocation <> ''
ORDER BY ScanID;
GO

-- ---------------------------------------------------------------------------
-- 7. Template literal rows — these still contain [Field] tokens.
--    Their root is now corrected (/Company/...) but the tokens remain.
--    Review and fix manually.
-- ---------------------------------------------------------------------------
SELECT
    CaseDocumentID,
    CaseID,
    DocumentFileName
FROM tblCaseDocuments
WHERE DocumentFileName LIKE '%[[]%'
ORDER BY CaseDocumentID;
GO
