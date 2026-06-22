SELECT tblCase.CaseID, Sum([Trust Account].Debit) AS SumOfDebit
FROM tblCase INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID
GROUP BY tblCase.CaseID, [Trust Account].DepCleared
HAVING ((([Trust Account].DepCleared)=Yes));