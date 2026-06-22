SELECT Sum(qryTTAmountHours_SUM_byAtty.SumOfTime_) AS SumOfSumOfTime_, qryTTAmountHours_SUM_byAtty.Tatty, [TB Time Keeping].CaseID
FROM qryTTAmountHours_SUM_byAtty LEFT JOIN [TB Time Keeping] ON qryTTAmountHours_SUM_byAtty.Bill_ID = [TB Time Keeping].Bill_ID
GROUP BY qryTTAmountHours_SUM_byAtty.Tatty, [TB Time Keeping].CaseID;