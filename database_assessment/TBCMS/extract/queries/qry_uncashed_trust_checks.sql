SELECT [Last_Name] & ", " & [First_Name] AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Trust Account].CheckCashed, [Trust Account].DepCleared, [Trust Account].CaseID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckNumber
FROM TblCase INNER JOIN [Trust Account] ON TblCase.CaseID = [Trust Account].CaseID
WHERE ((([Trust Account].CheckCashed)=Yes)) OR ((([Trust Account].DepCleared)=Yes))
ORDER BY [Last_Name] & ", " & [First_Name];