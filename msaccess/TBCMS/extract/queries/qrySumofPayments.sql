SELECT [Matter and AR].CaseID, Sum([Matter and AR].Payment) AS SumOfPayment
FROM [Matter and AR]
WHERE ((([Matter and AR].Pay_Outlay) Like "*payment*"))
GROUP BY [Matter and AR].CaseID;