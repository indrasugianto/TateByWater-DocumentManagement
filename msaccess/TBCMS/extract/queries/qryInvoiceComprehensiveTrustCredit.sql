SELECT [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Credit, [Trust Account].OrderNr, [Trust Account].CaseID, [TB Time Keeping].[BilL Closed Date]
FROM [TB Time Keeping] LEFT JOIN [Trust Account] ON [TB Time Keeping].CaseID = [Trust Account].CaseID
WHERE ((([Trust Account].TMatter) Not Like "*Earned*" And ([Trust Account].TMatter) Not Like "*Reimb*" And ([Trust Account].TMatter) Not Like "*Refund*") AND (([Trust Account].Credit)>0) AND (([Trust Account].CaseID)=9966))
ORDER BY [Trust Account].TDate;