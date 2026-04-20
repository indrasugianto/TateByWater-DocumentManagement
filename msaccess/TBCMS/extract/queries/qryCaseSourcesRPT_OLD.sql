SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.CaseOpenDate, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS CaseNo, tblCase.yr, Disposition.[Total Earned Fee], tblCase.Clsdate, tblCase.Closed
FROM tblCase INNER JOIN Disposition ON tblCase.CaseID = Disposition.CaseID
WHERE (((tblCase.Closed)=Yes));