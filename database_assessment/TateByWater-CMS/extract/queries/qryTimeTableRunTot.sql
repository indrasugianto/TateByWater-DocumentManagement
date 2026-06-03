SELECT qryTTAmount.Time_ID, qryTTAmount.Bill_ID, qryTTAmount.Time_, qryTTAmount.Rate, qryTTAmount.Amount, (Select sum(amount) from qryTTAmount as TT where time_ID<= qryTTAmount.Time_ID and bill_ID=GetBill_ID()) AS Run_total
FROM qryTTAmount
WHERE (((qryTTAmount.Bill_ID)=GetBill_ID()));