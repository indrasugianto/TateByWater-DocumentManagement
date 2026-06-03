SELECT Sum(qryTrustCostsExpended.CostBalance) AS SumOfCostBalance, qryTrustCostsExpended.CaseID
FROM qryTrustCostsExpended
GROUP BY qryTrustCostsExpended.CaseID;