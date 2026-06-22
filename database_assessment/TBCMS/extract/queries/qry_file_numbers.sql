SELECT tblCase.CaseID, Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNo
FROM tblCase;