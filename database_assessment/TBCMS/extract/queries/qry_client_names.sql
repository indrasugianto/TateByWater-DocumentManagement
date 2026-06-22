SELECT tblCase.CaseID, [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS Name
FROM tblCase
ORDER BY [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_");