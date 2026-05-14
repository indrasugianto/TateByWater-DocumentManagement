-- =============================================================================
-- STEP_2_VERIFY.sql
--
-- PURPOSE  : Post-migration verification. Run after STEP_1_UPDATE is committed.
--            All counts in the "Leftover offenders" block must be 0 (per table)
--            for the migration to be considered clean.
--
--            Manual-triage rows in dbo.tblScans_ManualTriage and
--            dbo.TBIntakes_ManualTriage are excluded from the leftover checks —
--            they are intentionally untouched and must be hand-fixed separately.
--
-- SERVER   : awsql2022dev (test environment for Phases 0–6)
-- DATABASE : TateByWater
-- =============================================================================

USE TateByWater;
GO

-- ---------------------------------------------------------------------------
-- 1. Leftover offenders — Dropbox paths must contain none of: '\', 'S:',
--    leading '#', trailing '#', 'file:///', '\\TBF-SRVR12\', or '%20'.
--    Quarantined rows are filtered out so they do not register as offenders.
--    Expected: every numeric column = 0 on every row of this query.
-- ---------------------------------------------------------------------------

SELECT 'tblCaseDocuments'                                                                 AS TableName,
       SUM(CASE WHEN DocumentFileName LIKE '%\%'                          THEN 1 ELSE 0 END) AS HasBackslash,
       SUM(CASE WHEN DocumentFileName LIKE '%S:%'                         THEN 1 ELSE 0 END) AS HasSColon,
       SUM(CASE WHEN LEFT(DocumentFileName,1) = '#'                       THEN 1 ELSE 0 END) AS LeadingHash,
       SUM(CASE WHEN RIGHT(DocumentFileName,1) = '#'                      THEN 1 ELSE 0 END) AS TrailingHash,
       SUM(CASE WHEN DocumentFileName LIKE '%file:///%'                   THEN 1 ELSE 0 END) AS HasFileURL,
       SUM(CASE WHEN DocumentFileName LIKE '%\\TBF-SRVR12\%'              THEN 1 ELSE 0 END) AS HasLegacyUNC,
       SUM(CASE WHEN DocumentFileName LIKE '%[%]20%' COLLATE Latin1_General_BIN THEN 1 ELSE 0 END) AS HasUrlEncodedSpace
FROM dbo.tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> '';

SELECT 'tblScans'                                                                           AS TableName,
       SUM(CASE WHEN s.ScanLocation LIKE '%\%'                            THEN 1 ELSE 0 END) AS HasBackslash,
       SUM(CASE WHEN s.ScanLocation LIKE '%S:%'                           THEN 1 ELSE 0 END) AS HasSColon,
       SUM(CASE WHEN LEFT(s.ScanLocation,1) = '#'                         THEN 1 ELSE 0 END) AS LeadingHash,
       SUM(CASE WHEN RIGHT(s.ScanLocation,1) = '#'                        THEN 1 ELSE 0 END) AS TrailingHash,
       SUM(CASE WHEN s.ScanLocation LIKE '%file:///%'                     THEN 1 ELSE 0 END) AS HasFileURL,
       SUM(CASE WHEN s.ScanLocation LIKE '%\\TBF-SRVR12\%'                THEN 1 ELSE 0 END) AS HasLegacyUNC,
       SUM(CASE WHEN s.ScanLocation LIKE '%[%]20%' COLLATE Latin1_General_BIN THEN 1 ELSE 0 END) AS HasUrlEncodedSpace
FROM dbo.tblScans s
WHERE s.ScanLocation IS NOT NULL
  AND s.ScanLocation <> ''
  AND NOT EXISTS (SELECT 1 FROM dbo.tblScans_ManualTriage q WHERE q.ScansID = s.ScansID);

SELECT 'TB Intakes'                                                                         AS TableName,
       SUM(CASE WHEN [Scan Location GI] LIKE '%\%'                        THEN 1 ELSE 0 END) AS HasBackslash,
       SUM(CASE WHEN [Scan Location GI] LIKE '%S:%'                       THEN 1 ELSE 0 END) AS HasSColon,
       SUM(CASE WHEN LEFT([Scan Location GI],1) = '#'                     THEN 1 ELSE 0 END) AS LeadingHash,
       SUM(CASE WHEN RIGHT([Scan Location GI],1) = '#'                    THEN 1 ELSE 0 END) AS TrailingHash,
       SUM(CASE WHEN [Scan Location GI] LIKE '%file:///%'                 THEN 1 ELSE 0 END) AS HasFileURL,
       SUM(CASE WHEN [Scan Location GI] LIKE '%\\TBF-SRVR12\%'            THEN 1 ELSE 0 END) AS HasLegacyUNC,
       SUM(CASE WHEN [Scan Location GI] LIKE '%[%]20%' COLLATE Latin1_General_BIN THEN 1 ELSE 0 END) AS HasUrlEncodedSpace
FROM [TB Intakes]
WHERE [Scan Location GI] IS NOT NULL
  AND [Scan Location GI] <> ''
  AND NOT EXISTS (
      SELECT 1 FROM dbo.TBIntakes_ManualTriage q
       WHERE q.OriginalValue = [TB Intakes].[Scan Location GI]);
GO

-- ---------------------------------------------------------------------------
-- 2. Detail listing — which rows still match any offender pattern?
--    Expected: zero rows returned across all three queries below.
-- ---------------------------------------------------------------------------

SELECT CaseDocumentID, DocumentFileName
FROM dbo.tblCaseDocuments
WHERE DocumentFileName IS NOT NULL
  AND DocumentFileName <> ''
  AND (   DocumentFileName LIKE '%\%'
       OR DocumentFileName LIKE '%S:%'
       OR LEFT(DocumentFileName,1) = '#'
       OR RIGHT(DocumentFileName,1) = '#'
       OR DocumentFileName LIKE '%file:///%'
       OR DocumentFileName LIKE '%\\TBF-SRVR12\%'
       OR DocumentFileName LIKE '%[%]20%' COLLATE Latin1_General_BIN);
GO

SELECT s.ScansID, s.ScanLocation
FROM dbo.tblScans s
WHERE s.ScanLocation IS NOT NULL
  AND s.ScanLocation <> ''
  AND NOT EXISTS (SELECT 1 FROM dbo.tblScans_ManualTriage q WHERE q.ScansID = s.ScansID)
  AND (   s.ScanLocation LIKE '%\%'
       OR s.ScanLocation LIKE '%S:%'
       OR LEFT(s.ScanLocation,1) = '#'
       OR RIGHT(s.ScanLocation,1) = '#'
       OR s.ScanLocation LIKE '%file:///%'
       OR s.ScanLocation LIKE '%\\TBF-SRVR12\%'
       OR s.ScanLocation LIKE '%[%]20%' COLLATE Latin1_General_BIN);
GO

SELECT [GI Last Name], [GI First Name], [GI Date], [Scan Location GI]
FROM [TB Intakes]
WHERE [Scan Location GI] IS NOT NULL
  AND [Scan Location GI] <> ''
  AND NOT EXISTS (
      SELECT 1 FROM dbo.TBIntakes_ManualTriage q
       WHERE q.OriginalValue = [TB Intakes].[Scan Location GI])
  AND (   [Scan Location GI] LIKE '%\%'
       OR [Scan Location GI] LIKE '%S:%'
       OR LEFT([Scan Location GI],1) = '#'
       OR RIGHT([Scan Location GI],1) = '#'
       OR [Scan Location GI] LIKE '%file:///%'
       OR [Scan Location GI] LIKE '%\\TBF-SRVR12\%'
       OR [Scan Location GI] LIKE '%[%]20%' COLLATE Latin1_General_BIN);
GO

-- ---------------------------------------------------------------------------
-- 3. Path prefix distribution — should be 100% '/Company/...' for non-NULL,
--    non-triaged rows on every table.
-- ---------------------------------------------------------------------------

SELECT 'tblCaseDocuments' AS TableName, LEFT(DocumentFileName, 30) AS PathPrefix, COUNT(*) AS RowCount
FROM dbo.tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> ''
GROUP BY LEFT(DocumentFileName, 30)
ORDER BY COUNT(*) DESC;
GO

SELECT 'tblScans' AS TableName, LEFT(s.ScanLocation, 30) AS PathPrefix, COUNT(*) AS RowCount
FROM dbo.tblScans s
WHERE s.ScanLocation IS NOT NULL AND s.ScanLocation <> ''
  AND NOT EXISTS (SELECT 1 FROM dbo.tblScans_ManualTriage q WHERE q.ScansID = s.ScansID)
GROUP BY LEFT(s.ScanLocation, 30)
ORDER BY COUNT(*) DESC;
GO

SELECT 'TB Intakes' AS TableName, LEFT([Scan Location GI], 30) AS PathPrefix, COUNT(*) AS RowCount
FROM [TB Intakes]
WHERE [Scan Location GI] IS NOT NULL AND [Scan Location GI] <> ''
  AND NOT EXISTS (
      SELECT 1 FROM dbo.TBIntakes_ManualTriage q
       WHERE q.OriginalValue = [TB Intakes].[Scan Location GI])
GROUP BY LEFT([Scan Location GI], 30)
ORDER BY COUNT(*) DESC;
GO

-- ---------------------------------------------------------------------------
-- 4. Sample of migrated paths for human review (20 rows per table)
-- ---------------------------------------------------------------------------

SELECT TOP 20
    CaseDocumentID,
    CaseID,
    DocumentFileName
FROM dbo.tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> ''
ORDER BY CaseDocumentID;
GO

SELECT TOP 20
    ScansID,
    CaseID,
    ScanLocation
FROM dbo.tblScans
WHERE ScanLocation IS NOT NULL AND ScanLocation <> ''
ORDER BY ScansID;
GO

SELECT TOP 20
    [GI Last Name],
    [GI First Name],
    [GI Date],
    [Scan Location GI]
FROM [TB Intakes]
WHERE [Scan Location GI] IS NOT NULL AND [Scan Location GI] <> ''
ORDER BY [GI Last Name];
GO

-- ---------------------------------------------------------------------------
-- 5. Defective rows that survived migration with their root corrected but
--    their content still needing manual repair (analysed in Phase 1b):
--      (a) 9 rows with unresolved template tokens (e.g., '[Case_Letter]')
--      (b) 4 rows with truncated '.[df' filenames
--    These 13 rows match LIKE '%[[]%'. Their Dropbox root is /Company/, but
--    the file name is still defective — they cannot resolve in Dropbox.
-- ---------------------------------------------------------------------------

SELECT
    CaseDocumentID,
    CaseID,
    DocumentFileName,
    CASE
        WHEN DocumentFileName LIKE '%[[]Case_Letter[]]%' THEN 'unresolved template'
        WHEN DocumentFileName LIKE '%.[[]df'             THEN 'truncated .[df filename'
        ELSE                                                  'other bracket defect'
    END AS DefectType
FROM dbo.tblCaseDocuments
WHERE DocumentFileName LIKE '%[[]%'
ORDER BY CaseDocumentID;
GO

-- ---------------------------------------------------------------------------
-- 6. Manual-triage tables — rows that STEP_1_UPDATE.sql did NOT rewrite.
--    Review row-by-row and either fix or delete before Phase 7.
-- ---------------------------------------------------------------------------

SELECT * FROM dbo.tblScans_ManualTriage     ORDER BY ScansID;
GO

SELECT * FROM dbo.TBIntakes_ManualTriage    ORDER BY TriageID;
GO
