SELECT TblCase.CaseID, [last_name] & ", " & [first_name] AS Name, TblCase.Closed, TblCase.Scan, TblCase.[Scan Location], TblCase.ScanNotAvail
FROM TblCase
GROUP BY TblCase.CaseID, [last_name] & ", " & [first_name], TblCase.Closed, TblCase.Scan, TblCase.[Scan Location], TblCase.ScanNotAvail, TblCase.Last_Name
HAVING (((TblCase.Closed)=Yes) AND ((TblCase.Scan)=No) AND ((TblCase.ScanNotAvail)=No) AND ((TblCase.Last_Name) Is Not Null)) OR (((TblCase.[Scan Location])="No"))
ORDER BY [last_name] & ", " & [first_name];