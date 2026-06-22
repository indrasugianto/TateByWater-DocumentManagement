SELECT TblCase.CaseID, [last_name] & ", " & [first_name] AS Name
FROM TblCase
GROUP BY TblCase.CaseID, [last_name] & ", " & [first_name], TblCase.Last_Name
HAVING (((TblCase.Last_Name) Is Not Null))
ORDER BY [last_name] & ", " & [first_name];