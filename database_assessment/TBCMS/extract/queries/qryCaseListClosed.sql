SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, TblCase.Extended_Ledger, TblCase.Court, TblCase.Matter_type, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.[Scan Location], TblCase.HandlingAtty_Case, TblCase.Closed, tblDropD.CodeVal, TblCase.Clsdate, TblCase.ParaLegal
FROM TblCase LEFT JOIN tblDropD ON TblCase.Case_Letter = tblDropD.Code
WHERE (((TblCase.Closed)=Yes))
ORDER BY [Last_Name] & ", " & [First_Name];