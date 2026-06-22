SELECT [Last_Name] & ", " & [First_Name] AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Matter and AR].*, TblCase.Orig_Atty, TblCase.Case_Letter, tblDropD.CodeVal, [Matter and AR].Date2
FROM (TblCase LEFT JOIN tblDropD ON TblCase.Case_Letter = tblDropD.Code) INNER JOIN [Matter and AR] ON TblCase.CaseID = [Matter and AR].CaseID
WHERE ((([Matter and AR].FirmPrepaid)=Yes))
ORDER BY [Matter and AR].Date2 DESC;