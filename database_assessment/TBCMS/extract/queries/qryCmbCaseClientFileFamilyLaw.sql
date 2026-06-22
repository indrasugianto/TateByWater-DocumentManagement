SELECT TblCase.CaseID, [TblCase].Last_Name & ", " & [TblCase].First_Name AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber
FROM (TblCase RIGHT JOIN [Family Law - Divorce] ON TblCase.CaseID = [Family Law - Divorce].CaseID) LEFT JOIN [TB Time Keeping] ON TblCase.CaseID = [TB Time Keeping].CaseID
GROUP BY TblCase.CaseID, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_"), [TblCase].Last_Name & ", " & [TblCase].Last_Name, TblCase.First_Name, TblCase.Last_Name
ORDER BY TblCase.First_Name, TblCase.Last_Name;