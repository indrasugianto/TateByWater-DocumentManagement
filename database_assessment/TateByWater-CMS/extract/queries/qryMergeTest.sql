SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Court, tbl_CtCaseNumbers.CtNumber, tblCase.CaseID, tblCase.Closed
FROM tblCase INNER JOIN tbl_CtCaseNumbers ON tblCase.CaseID = tbl_CtCaseNumbers.CaseID
WHERE (((tblCase.Closed)=No));