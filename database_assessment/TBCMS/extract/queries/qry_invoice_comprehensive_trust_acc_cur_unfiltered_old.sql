SELECT [Trust Account].CaseID, [Trust Account].OrderNr, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, fncGetTABalanceWithCaseID([ordernr],[caseid]) AS Balance
FROM [Trust Account]
ORDER BY [Trust Account].OrderNr;