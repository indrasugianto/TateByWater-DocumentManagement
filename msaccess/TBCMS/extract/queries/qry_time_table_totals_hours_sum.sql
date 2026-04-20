SELECT qryTTAmount.Bill_ID, Sum(qryTTAmount.Amount) AS SumOfAmount
FROM qryTTAmount
GROUP BY qryTTAmount.Bill_ID;