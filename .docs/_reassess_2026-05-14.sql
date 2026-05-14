SET NOCOUNT ON;
USE TateByWater;

PRINT '=== A. Core row counts ===';
SELECT
    (SELECT COUNT(*) FROM tblCaseDocuments)      AS tblCaseDocuments,
    (SELECT COUNT(*) FROM tblScans)              AS tblScans,
    (SELECT COUNT(*) FROM tblCase)               AS tblCase,
    (SELECT COUNT(*) FROM tblDocumentTypes)      AS tblDocumentTypes,
    (SELECT COUNT(*) FROM tblDocumentRootDirectory) AS tblDocumentRootDirectory;

PRINT '=== B. tblCase coverage ===';
SELECT
    (SELECT COUNT(*) FROM tblCase) AS cases_total,
    (SELECT COUNT(*) FROM tblCase c
       WHERE EXISTS (SELECT 1 FROM tblCaseDocuments d WHERE d.CaseID = c.CaseID)) AS cases_with_docs,
    (SELECT COUNT(*) FROM tblCase c
       WHERE NOT EXISTS (SELECT 1 FROM tblCaseDocuments d WHERE d.CaseID = c.CaseID)) AS cases_without_docs;

PRINT '=== C. Distinct (CaseID, DocumentType) pairs and top-N outliers ===';
SELECT COUNT(*) AS distinct_pairs
FROM (SELECT DISTINCT CaseID, DocumentType FROM tblCaseDocuments) p;

SELECT TOP 5 CaseID, DocumentType, COUNT(*) AS row_count
FROM tblCaseDocuments
GROUP BY CaseID, DocumentType
ORDER BY COUNT(*) DESC;

PRINT '=== D. tblCaseDocuments path-prefix distribution (first 20 chars) ===';
SELECT LEFT(DocumentFileName, 20) AS PathPrefix, COUNT(*) AS rows_count
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> ''
GROUP BY LEFT(DocumentFileName, 20)
ORDER BY COUNT(*) DESC;

PRINT '=== E. tblCaseDocuments prefix summary (S vs hashS vs /Company vs other) ===';
SELECT
    COUNT(*)                                                                  AS total_rows,
    SUM(CASE WHEN DocumentFileName IS NULL OR DocumentFileName = '' THEN 1 ELSE 0 END) AS null_or_empty,
    SUM(CASE WHEN DocumentFileName LIKE 'S:\%'             THEN 1 ELSE 0 END) AS starts_S,
    SUM(CASE WHEN DocumentFileName LIKE '#S:\%'            THEN 1 ELSE 0 END) AS starts_hashS,
    SUM(CASE WHEN DocumentFileName LIKE '/Company/%'       THEN 1 ELSE 0 END) AS starts_Company,
    SUM(CASE WHEN DocumentFileName LIKE '#/Company/%'      THEN 1 ELSE 0 END) AS starts_hashCompany,
    SUM(CASE WHEN DocumentFileName IS NOT NULL
              AND DocumentFileName <> ''
              AND DocumentFileName NOT LIKE 'S:\%'
              AND DocumentFileName NOT LIKE '#S:\%'
              AND DocumentFileName NOT LIKE '/Company/%'
              AND DocumentFileName NOT LIKE '#/Company/%' THEN 1 ELSE 0 END) AS other
FROM tblCaseDocuments;

PRINT '=== F. tblCaseDocuments unresolved template literals ([Field] markers) ===';
SELECT COUNT(*) AS rows_with_brackets
FROM tblCaseDocuments
WHERE DocumentFileName LIKE '%[[]%';

PRINT '=== G. tblCaseDocuments residual backslashes (should be 0 post-migration) ===';
SELECT COUNT(*) AS rows_with_backslash
FROM tblCaseDocuments
WHERE DocumentFileName LIKE '%\%';

PRINT '=== H. tblScans path-prefix distribution (first 20 chars) ===';
SELECT LEFT(ScanLocation, 20) AS PathPrefix, COUNT(*) AS rows_count
FROM tblScans
WHERE ScanLocation IS NOT NULL AND ScanLocation <> ''
GROUP BY LEFT(ScanLocation, 20)
ORDER BY COUNT(*) DESC;

PRINT '=== I. tblScans prefix summary ===';
SELECT
    COUNT(*)                                                              AS total_rows,
    SUM(CASE WHEN ScanLocation IS NULL OR ScanLocation = '' THEN 1 ELSE 0 END) AS null_or_empty,
    SUM(CASE WHEN ScanLocation LIKE 'S:\%'           THEN 1 ELSE 0 END) AS starts_S,
    SUM(CASE WHEN ScanLocation LIKE '#S:\%'          THEN 1 ELSE 0 END) AS starts_hashS,
    SUM(CASE WHEN ScanLocation LIKE '/Company/%'     THEN 1 ELSE 0 END) AS starts_Company,
    SUM(CASE WHEN ScanLocation LIKE '#/Company/%'    THEN 1 ELSE 0 END) AS starts_hashCompany,
    SUM(CASE WHEN ScanLocation IS NOT NULL
              AND ScanLocation <> ''
              AND ScanLocation NOT LIKE 'S:\%'
              AND ScanLocation NOT LIKE '#S:\%'
              AND ScanLocation NOT LIKE '/Company/%'
              AND ScanLocation NOT LIKE '#/Company/%' THEN 1 ELSE 0 END) AS other
FROM tblScans;

PRINT '=== J. tblScans residual backslashes ===';
SELECT COUNT(*) AS rows_with_backslash
FROM tblScans
WHERE ScanLocation LIKE '%\%';

PRINT '=== K. tblScans TypeofScan null fraction ===';
SELECT
    COUNT(*) AS total,
    SUM(CASE WHEN TypeofScan IS NULL THEN 1 ELSE 0 END) AS null_count,
    SUM(CASE WHEN TypeofScan IS NULL THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*),0) AS null_pct
FROM tblScans;

PRINT '=== L. tblDocumentTypes — total and visibility breakdown ===';
SELECT
    COUNT(*) AS total,
    SUM(CASE WHEN IsVisible = 1 THEN 1 ELSE 0 END) AS visible,
    SUM(CASE WHEN IsVisible = 0 OR IsVisible IS NULL THEN 1 ELSE 0 END) AS hidden
FROM tblDocumentTypes;

SELECT DocumentTypeID, DocumentType, DocumentFolder, IsVisible
FROM tblDocumentTypes
ORDER BY IsVisible DESC, DocumentTypeID;

PRINT '=== M. Dropbox migration-related new tables — existence check ===';
SELECT
    OBJECT_ID('dbo.tblDropboxConfig')             AS tblDropboxConfig,
    OBJECT_ID('dbo.tblDropboxRootConfig')         AS tblDropboxRootConfig,
    OBJECT_ID('dbo.tblDropboxRevocationList')     AS tblDropboxRevocationList,
    OBJECT_ID('dbo.tblDropboxAuditLog')           AS tblDropboxAuditLog,
    OBJECT_ID('dbo.tblDropboxVerificationReport') AS tblDropboxVerificationReport;

PRINT '=== N. Pre-migration snapshot tables (look for any *PreDropbox* or backups) ===';
SELECT name, create_date
FROM sys.tables
WHERE name LIKE '%Dropbox%'
   OR name LIKE '%PreDropbox%'
   OR name LIKE '%_backup%'
   OR name LIKE 'tblCaseDocuments_%'
   OR name LIKE 'tblScans_%';

PRINT '=== O. Stored procedure inventory — Dropbox-related and document-related ===';
SELECT name, create_date, modify_date
FROM sys.procedures
WHERE name LIKE 'sp%Document%' OR name LIKE 'sp%Dropbox%' OR name LIKE 'sp%Scan%' OR name LIKE 'sp%CaseDocument%'
ORDER BY name;

PRINT '=== P. Sample of migrated tblCaseDocuments rows (10) ===';
SELECT TOP 10 CaseDocumentID, CaseID, DocumentType, DocumentFileName
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> ''
ORDER BY CaseDocumentID DESC;

PRINT '=== Q. Sample of migrated tblScans rows (10) ===';
SELECT TOP 10 ScansID, ScanLocation
FROM tblScans
WHERE ScanLocation IS NOT NULL AND ScanLocation <> ''
ORDER BY ScansID DESC;

PRINT '=== R. Path-length distribution (Dropbox ~260 char limit awareness) ===';
SELECT
    MIN(LEN(DocumentFileName)) AS min_len,
    MAX(LEN(DocumentFileName)) AS max_len,
    AVG(LEN(DocumentFileName)) AS avg_len,
    SUM(CASE WHEN LEN(DocumentFileName) > 260 THEN 1 ELSE 0 END) AS over_260,
    SUM(CASE WHEN LEN(DocumentFileName) > 200 THEN 1 ELSE 0 END) AS over_200
FROM tblCaseDocuments
WHERE DocumentFileName IS NOT NULL AND DocumentFileName <> '';

PRINT '=== S. tblDocumentRootDirectory current contents (template source of truth) ===';
SELECT * FROM tblDocumentRootDirectory;

PRINT '=== T. tblDocumentTypes structure check — list columns ===';
SELECT COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'tblDocumentTypes'
ORDER BY ORDINAL_POSITION;
