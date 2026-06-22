SELECT [Trust Account].CaseID, [Trust Account].OrderNr, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, fncGetTABalanceWithCaseID([ordernr],[caseid]) AS Balance
FROM [Trust Account]
WHERE ((([Trust Account].TMatter) Not Like "*Earned*" And ([Trust Account].TMatter) Not Like "*Reimb*" And ([Trust Account].TMatter) Not Like "*Wire*" And ([Trust Account].TMatter) Not Like "*refund*"))
ORDER BY [Trust Account].OrderNr;