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
-- 9. tblScans path prefix categorisation
--    Mirrors the per-pass classification in STEP_1_UPDATE.sql. Each row should
--    fall into exactly one category. The OTHER bucket is quarantined.
-- ---------------------------------------------------------------------------
SELECT
    CASE
        WHEN ScanLocation IS NULL OR LTRIM(RTRIM(ScanLocation)) = '' THEN 'NULL/EMPTY'
        WHEN ScanLocation LIKE '#S:\%'                  THEN 'B1 #S:\…#'
        WHEN ScanLocation LIKE 'S:\%'                   THEN 'B2 S:\… (displaytext#URL#)'
        WHEN ScanLocation LIKE '#file:///S:\%'          THEN 'B3 #file:///S:\…#'
        WHEN ScanLocation LIKE '#?S:\%'                 THEN 'B4 #?S:\…#'
        WHEN ScanLocation LIKE '#file:///\\TBF-SRVR12\%' THEN 'B5 #file:///\\TBF-SRVR12\…#'
        WHEN ScanLocation LIKE '#\\TBF-SRVR12\%'        THEN 'B6 #\\TBF-SRVR12\…#'
        ELSE 'OTHER (quarantine)'
    END                                                         AS Category,
    COUNT(*)                                                    AS RowCount
FROM tblScans
GROUP BY
    CASE
        WHEN ScanLocation IS NULL OR LTRIM(RTRIM(ScanLocation)) = '' THEN 'NULL/EMPTY'
        WHEN ScanLocation LIKE '#S:\%'                  THEN 'B1 #S:\…#'
        WHEN ScanLocation LIKE 'S:\%'                   THEN 'B2 S:\… (displaytext#URL#)'
        WHEN ScanLocation LIKE '#file:///S:\%'          THEN 'B3 #file:///S:\…#'
        WHEN ScanLocation LIKE '#?S:\%'                 THEN 'B4 #?S:\…#'
        WHEN ScanLocation LIKE '#file:///\\TBF-SRVR12\%' THEN 'B5 #file:///\\TBF-SRVR12\…#'
        WHEN ScanLocation LIKE '#\\TBF-SRVR12\%'        THEN 'B6 #\\TBF-SRVR12\…#'
        ELSE 'OTHER (quarantine)'
    END
ORDER BY COUNT(*) DESC;
GO

-- ---------------------------------------------------------------------------
-- 10. tblScans OTHER rows — these will be copied to dbo.tblScans_ManualTriage
--     by STEP_1_UPDATE.sql, not rewritten. Review before commit.
-- ---------------------------------------------------------------------------
SELECT ScansID, ScanLocation
FROM tblScans
WHERE ScanLocation IS NOT NULL
  AND LTRIM(RTRIM(ScanLocation)) <> ''
  AND ScanLocation NOT LIKE '#S:\%'
  AND ScanLocation NOT LIKE 'S:\%'
  AND ScanLocation NOT LIKE '#file:///S:\%'
  AND ScanLocation NOT LIKE '#?S:\%'
  AND ScanLocation NOT LIKE '#file:///\\TBF-SRVR12\%'
  AND ScanLocation NOT LIKE '#\\TBF-SRVR12\%'
ORDER BY ScansID;
GO

-- ---------------------------------------------------------------------------
-- 11. [TB Intakes].[Scan Location GI] categorisation
--     STEP_1_UPDATE.sql Part C handles the four recoverable categories;
--     OTHER is quarantined to dbo.TBIntakes_ManualTriage.
-- ---------------------------------------------------------------------------
SELECT
    CASE
        WHEN [Scan Location GI] IS NULL OR LTRIM(RTRIM([Scan Location GI])) = '' THEN 'NULL/EMPTY'
        WHEN [Scan Location GI] LIKE 'S:\%'             THEN 'C1 S:\…'
        WHEN [Scan Location GI] LIKE '#S:\%'            THEN 'C2 #S:\…#'
        WHEN [Scan Location GI] LIKE '#file:///S:\%'    THEN 'C3 #file:///S:\…#'
        WHEN [Scan Location GI] LIKE '?S:\%'            THEN 'C4 ?S:\…'
        ELSE 'OTHER (quarantine)'
    END                                                         AS Category,
    COUNT(*)                                                    AS RowCount
FROM [TB Intakes]
GROUP BY
    CASE
        WHEN [Scan Location GI] IS NULL OR LTRIM(RTRIM([Scan Location GI])) = '' THEN 'NULL/EMPTY'
        WHEN [Scan Location GI] LIKE 'S:\%'             THEN 'C1 S:\…'
        WHEN [Scan Location GI] LIKE '#S:\%'            THEN 'C2 #S:\…#'
        WHEN [Scan Location GI] LIKE '#file:///S:\%'    THEN 'C3 #file:///S:\…#'
        WHEN [Scan Location GI] LIKE '?S:\%'            THEN 'C4 ?S:\…'
        ELSE 'OTHER (quarantine)'
    END
ORDER BY COUNT(*) DESC;
GO

-- ---------------------------------------------------------------------------
-- 12. [TB Intakes] OTHER rows — review before commit.
-- ---------------------------------------------------------------------------
SELECT [GI Last Name], [GI First Name], [GI Date], [Scan Location GI]
FROM [TB Intakes]
WHERE [Scan Location GI] IS NOT NULL
  AND LTRIM(RTRIM([Scan Location GI])) <> ''
  AND [Scan Location GI] NOT LIKE 'S:\%'
  AND [Scan Location GI] NOT LIKE '#S:\%'
  AND [Scan Location GI] NOT LIKE '#file:///S:\%'
  AND [Scan Location GI] NOT LIKE '?S:\%'
ORDER BY [GI Last Name];
GO
