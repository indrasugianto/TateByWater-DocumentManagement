SELECT vwTrustAccountBalanceTotals.CaseID, Max(vwTrustAccountBalanceTotals.SumOfBalance) AS SumOfBalance
FROM vwTrustAccountBalanceTotals
GROUP BY vwTrustAccountBalanceTotals.CaseID;