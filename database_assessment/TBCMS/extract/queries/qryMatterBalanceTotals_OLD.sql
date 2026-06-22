SELECT qryMatter.CaseID, Sum(qryMatter.Balance) AS SumOfBalance
FROM qryMatter
GROUP BY qryMatter.CaseID;