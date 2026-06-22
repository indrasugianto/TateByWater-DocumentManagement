SELECT [Trust Account].CaseID, Sum([Trust Account].Debit) AS SumOfDebit, Sum([Trust Account].Credit) AS SumOfCredit, Sum(Nz([debit],0)-Nz([Credit],0)) AS Balance, [Trust Account].OrderNr, [Trust Account].TrustAccountID
FROM [Trust Account]
GROUP BY [Trust Account].CaseID, [Trust Account].OrderNr, [Trust Account].TrustAccountID
ORDER BY [Trust Account].CaseID, [Trust Account].OrderNr;