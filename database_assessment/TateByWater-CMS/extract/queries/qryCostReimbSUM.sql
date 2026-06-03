SELECT vwCostReimbSUM.CaseID, vwCostReimbSUM.SumOfCredit
FROM vwCostReimbSUM
GROUP BY vwCostReimbSUM.CaseID, vwCostReimbSUM.SumOfCredit;