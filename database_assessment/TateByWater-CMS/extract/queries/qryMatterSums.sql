SELECT qryMatter.CaseID, Sum(qryMatter.SumOfCharge) AS SumOfSumOfCharge, Sum(qryMatter.SumOfPayment) AS SumOfSumOfPayment
FROM qryMatter
GROUP BY qryMatter.CaseID;