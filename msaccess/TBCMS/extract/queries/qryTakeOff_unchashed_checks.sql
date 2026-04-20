SELECT tblCase.CaseID, Sum([Trust Account].Credit) AS SumOfCredit
FROM tblCase INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID
GROUP BY tblCase.CaseID, [Trust Account].CheckCashed
HAVING ((([Trust Account].CheckCashed)=Yes));