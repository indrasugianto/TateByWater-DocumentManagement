SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.Closed, TblCase.Clsdate, TblCase.ARTrustZero, TblCase.Retainer, TblCase.FamilyLaw
FROM TblCase
WHERE (((TblCase.Closed)=No) AND ((TblCase.FamilyLaw)=Yes))
ORDER BY TblCase.Number_;