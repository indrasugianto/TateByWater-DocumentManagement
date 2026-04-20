SELECT [last_name] & ", " & [first_name] AS Name, TblCase.CaseID, TblCase.Closed
FROM TblCase
GROUP BY [last_name] & ", " & [first_name], TblCase.CaseID, TblCase.Closed, TblCase.Last_Name
HAVING (((TblCase.Closed)=No) AND ((TblCase.Last_Name) Is Not Null))
ORDER BY [last_name] & ", " & [first_name];