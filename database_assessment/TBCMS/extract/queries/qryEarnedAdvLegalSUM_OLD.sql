SELECT qryEarnedAdvLegal.CaseID, Sum(qryEarnedAdvLegal.Credit) AS SumOfCredit
FROM qryEarnedAdvLegal
GROUP BY qryEarnedAdvLegal.CaseID;