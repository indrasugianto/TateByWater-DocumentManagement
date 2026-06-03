SELECT TblCase.CaseID, [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS Name, TblCase.Closed
FROM TblCase
GROUP BY TblCase.CaseID, [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_"), TblCase.Closed, TblCase.Last_Name
HAVING (((TblCase.Closed)=No) AND ((TblCase.Last_Name) Is Not Null))
ORDER BY [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_");