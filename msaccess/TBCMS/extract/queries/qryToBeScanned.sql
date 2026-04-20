SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.[Scan Location], TblCase.HandlingAtty_Case, TblCase.Closed, TblCase.Scan, TblCase.Clsdate, TblCase.ScanNotAvail
FROM TblCase
WHERE (((TblCase.[Scan Location]) Is Null) AND ((TblCase.Closed)=Yes) AND ((TblCase.ScanNotAvail)=No))
ORDER BY TblCase.Number_;