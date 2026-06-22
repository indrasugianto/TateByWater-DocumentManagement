SELECT qryTrustAccount.CaseID, Sum(qryTrustAccount.Balance) AS SumOfBalance
FROM qryTrustAccount
GROUP BY qryTrustAccount.CaseID;