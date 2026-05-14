-- =============================================================================
-- FIX_DocumentTypes_NamingRule_Typo.sql
--
-- PURPOSE  : Phase 1b data-quality remediation. dbo.tblDocumentTypes row
--            (DocumentTypeID = 30, DocumentType = 'General') has the typo
--            '(customeruserentry)' (extra 'e') in its DocumentNamingRule.
--            Every other row uses the canonical '(customuserentry)' token,
--            which is the placeholder the path tokenizer (fnGetListOfWords)
--            recognises. As-is, filenames generated for the 'General'
--            document type contain a literal '(customeruserentry)' segment.
--
-- BEFORE   : Verify the typo still exists. Idempotent guard below skips the
--            update if the row already reads '(customuserentry)'.
--
-- AFTER    : Re-run STEP_0_ANALYZE-style sanity check on tblDocumentTypes:
--              SELECT DocumentTypeID, DocumentNamingRule
--              FROM dbo.tblDocumentTypes WHERE DocumentTypeID = 30;
--            Expected: DocumentNamingRule = '[Last_Name] (customuserentry)'.
--
-- SCOPE    : Phases 0–6 → run against awsql2022dev/TateByWater (test DB).
--            Phase 7    → re-run against the production DB as part of the
--                         Phase 1b production sync step.
--
-- SERVER   : awsql2022dev (test) / production SQL host (Phase 7)
-- DATABASE : TateByWater
-- =============================================================================

USE TateByWater;
GO

BEGIN TRANSACTION FixDocumentTypesTypo;

DECLARE @Before NVARCHAR(255), @After NVARCHAR(255), @Rows INT;

SELECT @Before = DocumentNamingRule
FROM dbo.tblDocumentTypes
WHERE DocumentTypeID = 30;

UPDATE dbo.tblDocumentTypes
SET DocumentNamingRule = REPLACE(DocumentNamingRule, '(customeruserentry)', '(customuserentry)')
WHERE DocumentTypeID  = 30
  AND DocumentNamingRule LIKE '%(customeruserentry)%';

SET @Rows = @@ROWCOUNT;

SELECT @After = DocumentNamingRule
FROM dbo.tblDocumentTypes
WHERE DocumentTypeID = 30;

SELECT
    @Rows                       AS RowsUpdated,
    @Before                     AS NamingRule_Before,
    @After                      AS NamingRule_After,
    CASE WHEN @After LIKE '%(customuserentry)%' AND @After NOT LIKE '%(customeruserentry)%'
         THEN 'OK — typo removed; canonical token in place'
         ELSE 'CHECK — review row 30 by hand' END AS Status;

-- ---------------------------------------------------------------------------
-- Sanity: confirm no other row carries the typo (defense in depth).
-- Expected: 0 rows.
-- ---------------------------------------------------------------------------
SELECT DocumentTypeID, DocumentType, DocumentNamingRule
FROM dbo.tblDocumentTypes
WHERE DocumentNamingRule LIKE '%(customeruserentry)%';

-- ---------------------------------------------------------------------------
-- DECISION POINT
--   If RowsUpdated = 1 (or 0 on idempotent re-run) and Status = OK →
--       COMMIT TRANSACTION FixDocumentTypesTypo;
--   Otherwise →
--       ROLLBACK TRANSACTION FixDocumentTypesTypo;
-- ---------------------------------------------------------------------------

-- COMMIT   TRANSACTION FixDocumentTypesTypo;
-- ROLLBACK TRANSACTION FixDocumentTypesTypo;
