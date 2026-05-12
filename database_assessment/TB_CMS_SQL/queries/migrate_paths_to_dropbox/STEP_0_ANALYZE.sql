-- =============================================================================
-- STEP_0_ANALYZE.sql
--
-- PURPOSE  : Inspect DocumentFileName (tblCaseDocuments) and ScanLocation
--            (tblScans) before running the path migration.
--            Run this first. No data is changed.
--
-- SERVER   : awsql2022dev
-- DATABASE : TateByWater
-- =============================================================================

USE TateByWater;
GO

-- ---------------------------------------------------------------------------
-- 1. Distinct path prefixes in tblCaseDocuments (first 20 chars)
--    Reveals all root patterns that the UPDATE must cover.
-- ---------------------------------------------------------------------------
SELECT
    LEFT(DocumentFileName, 20)  AS PathPrefix,
    COUNT(*)                    AS RowCount
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
GROUP BY LEFT(DocumentFileName, 20)
ORDER BY COUNT(*) DESC;
GO

-- ---------------------------------------------------------------------------
-- 2. Rows starting with #S:\ — show full sample to understand hyperlink format
-- ---------------------------------------------------------------------------
SELECT TOP 20
    CaseDocumentID,
    CaseID,
    DocumentFileName
FROM tblCaseDocuments
WHERE DocumentFileName LIKE '#S:\%'
ORDER BY CaseDocumentID;
GO

-- ---------------------------------------------------------------------------
-- 3. Rows with unresolved template literals ([...] placeholders)
-- ---------------------------------------------------------------------------
SELECT
    CaseDocumentID,
    CaseID,
    DocumentFileName
FROM tblCaseDocuments
WHERE DocumentFileName LIKE '%[[]%'  -- contains a literal [
ORDER BY CaseDocumentID;
GO

-- ---------------------------------------------------------------------------
-- 4. Rows with unexpected / already-migrated paths (not starting with S:\ or #S:\)
--    These should be zero before running the migration.
-- ---------------------------------------------------------------------------
SELECT
    CaseDocumentID,
    CaseID,
    DocumentFileName
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
  AND DocumentFileName NOT LIKE 'S:\%'
  AND DocumentFileName NOT LIKE '#S:\%'
ORDER BY CaseDocumentID;
GO

-- ---------------------------------------------------------------------------
-- 5. tblCaseDocuments row summary
-- ---------------------------------------------------------------------------
SELECT
    COUNT(*)                                        AS TotalRows,
    COUNT(CASE WHEN DocumentFileName IS NULL
               OR DocumentFileName = '' THEN 1 END) AS NullOrEmpty,
    COUNT(CASE WHEN DocumentFileName LIKE 'S:\%'    THEN 1 END) AS StartsWithS,
    COUNT(CASE WHEN DocumentFileName LIKE '#S:\%'   THEN 1 END) AS StartsWithHashS,
    COUNT(CASE WHEN DocumentFileName NOT LIKE 'S:\%'
               AND DocumentFileName NOT LIKE '#S:\%'
               AND DocumentFileName IS NOT NULL
               AND DocumentFileName <> ''           THEN 1 END) AS OtherPaths
FROM tblCaseDocuments;
GO

-- ---------------------------------------------------------------------------
-- 6. tblScans — inspect ScanLocation column (confirm column name and prefixes)
--    NOTE: If this query errors, the column name may differ.
--    Check: SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
--           WHERE TABLE_NAME = 'tblScans'
-- ---------------------------------------------------------------------------
SELECT
    LEFT(ScanLocation, 20)  AS PathPrefix,
    COUNT(*)                AS RowCount
FROM tblScans
WHERE ScanLocation IS NOT NULL
  AND ScanLocation <> ''
GROUP BY LEFT(ScanLocation, 20)
ORDER BY COUNT(*) DESC;
GO

-- ---------------------------------------------------------------------------
-- 7. tblScans row summary
-- ---------------------------------------------------------------------------
SELECT
    COUNT(*)                                     AS TotalRows,
    COUNT(CASE WHEN ScanLocation IS NULL
               OR ScanLocation = '' THEN 1 END)  AS NullOrEmpty,
    COUNT(CASE WHEN ScanLocation LIKE 'S:\%'     THEN 1 END) AS StartsWithS,
    COUNT(CASE WHEN ScanLocation LIKE '#S:\%'    THEN 1 END) AS StartsWithHashS,
    COUNT(CASE WHEN ScanLocation NOT LIKE 'S:\%'
               AND ScanLocation NOT LIKE '#S:\%'
               AND ScanLocation IS NOT NULL
               AND ScanLocation <> ''            THEN 1 END) AS OtherPaths
FROM tblScans;
GO

-- ---------------------------------------------------------------------------
-- 8. Preview: what DocumentFileName will look like AFTER migration
--    (dry run — no changes)
-- ---------------------------------------------------------------------------
SELECT TOP 30
    CaseDocumentID,
    DocumentFileName                                            AS Before,
    -- Strip leading # if present, replace S:\ with /Company/, flip slashes
    REPLACE(
        REPLACE(
            CASE WHEN LEFT(DocumentFileName, 1) = '#'
                 THEN SUBSTRING(DocumentFileName, 2, LEN(DocumentFileName))
                 ELSE DocumentFileName
            END,
            'S:\', '/Company/'
        ),
        '\', '/'
    )                                                           AS After
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
ORDER BY CaseDocumentID;
GO

-- ---------------------------------------------------------------------------
-- 9. Preview: tblScans ScanLocation after migration
-- ---------------------------------------------------------------------------
SELECT TOP 20
    ScanID,
    ScanLocation                                                AS Before,
    REPLACE(
        REPLACE(
            CASE WHEN LEFT(ScanLocation, 1) = '#'
                 THEN SUBSTRING(ScanLocation, 2, LEN(ScanLocation))
                 ELSE ScanLocation
            END,
            'S:\', '/Company/'
        ),
        '\', '/'
    )                                                           AS After
FROM tblScans
WHERE ScanLocation IS NOT NULL
  AND ScanLocation <> ''
ORDER BY ScanID;
GO
