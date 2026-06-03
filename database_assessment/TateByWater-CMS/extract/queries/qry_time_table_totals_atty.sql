SELECT qryTTAmountAtty.Bill_ID, qryTTAmountAtty.Tatty, Sum(qryTTAmountAtty.Amount) AS SumOfAmount
FROM qryTTAmountAtty
GROUP BY qryTTAmountAtty.Bill_ID, qryTTAmountAtty.Tatty;