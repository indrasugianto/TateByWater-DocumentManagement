SELECT tblCase.CaseID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Credit, [Trust Account].OrderNr
FROM tblCase INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID
WHERE ((([Trust Account].TMatter) Not Like "*Earned*" And ([Trust Account].TMatter) Not Like "*Reimb*" And ([Trust Account].TMatter) Not Like "*Refund*") AND (([Trust Account].Credit)>0))
ORDER BY [Trust Account].OrderNr;