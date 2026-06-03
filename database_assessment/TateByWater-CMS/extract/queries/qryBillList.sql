SELECT [TB Time Keeping].Bill_ID, [TB Time Keeping].IANumber, [TB Time Keeping].[BilL Closed Date], TblCase.CaseID
FROM TblCase INNER JOIN [TB Time Keeping] ON TblCase.CaseID = [TB Time Keeping].CaseID
ORDER BY [TB Time Keeping].Bill_ID, [TB Time Keeping].IANumber, TblCase.CaseID;