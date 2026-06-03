SELECT DISTINCT tblCase.CaseID, tblCase.Orig_Atty
FROM tblCase
GROUP BY tblCase.CaseID, tblCase.Orig_Atty;