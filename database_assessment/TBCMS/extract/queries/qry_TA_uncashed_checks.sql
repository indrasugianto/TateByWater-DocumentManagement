SELECT [Last_Name] & ", " & [First_Name] AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Trust Account].*
FROM TblCase INNER JOIN [Trust Account] ON TblCase.CaseID = [Trust Account].CaseID
WHERE ((([Trust Account].CheckCashed)=No))
ORDER BY [Last_Name] & ", " & [First_Name];