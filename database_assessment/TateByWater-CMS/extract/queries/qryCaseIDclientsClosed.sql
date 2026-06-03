SELECT TblCase.CaseID, [last_name] & ", " & [first_name] AS Name, TblCase.Closed
FROM TblCase
GROUP BY TblCase.CaseID, [last_name] & ", " & [first_name], TblCase.Closed, TblCase.Last_Name
HAVING (((TblCase.Closed)=Yes) AND ((TblCase.Last_Name) Is Not Null))
ORDER BY [last_name] & ", " & [first_name];