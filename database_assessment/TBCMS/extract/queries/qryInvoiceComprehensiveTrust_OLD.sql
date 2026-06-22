SELECT qryStmtTrustRPT.TrustAccountID, qryStmtTrustRPT.TDate, qryStmtTrustRPT.TMatter, qryStmtTrustRPT.Debit, qryStmtTrustRPT.Credit, tblCase.CaseID
FROM (qryStmtTrustRPT LEFT JOIN tblCase ON qryStmtTrustRPT.CaseID = tblCase.CaseID) LEFT JOIN qryTrustAccount ON qryStmtTrustRPT.TrustAccountID = qryTrustAccount.TrustAccountID
ORDER BY qryStmtTrustRPT.TDate;