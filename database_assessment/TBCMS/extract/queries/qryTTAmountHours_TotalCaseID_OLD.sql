SELECT Sum(qryTTAmountHours_SUM.SumOfTime_) AS SumOfSumOfTime_, [TB Time Keeping].CaseID
FROM qryTTAmountHours_SUM INNER JOIN [TB Time Keeping] ON qryTTAmountHours_SUM.Bill_ID = [TB Time Keeping].Bill_ID
GROUP BY [TB Time Keeping].CaseID;