SELECT tblCase.CaseID, Count([TB Time Keeping].IANumber) AS CountOfIANumber
FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID
GROUP BY tblCase.CaseID
HAVING (((Count([TB Time Keeping].IANumber)) Is Not Null));