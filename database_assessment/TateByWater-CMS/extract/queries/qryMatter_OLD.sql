SELECT [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].CaseID, Sum([Matter and AR].Charge) AS SumOfCharge, Sum([Matter and AR].Payment) AS SumOfPayment, Sum(Nz([Charge],0)-Nz([payment],0)) AS Balance, [Matter and AR].OrderNr
FROM [Matter and AR]
GROUP BY [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].CaseID, [Matter and AR].OrderNr
ORDER BY [Matter and AR].CaseID, [Matter and AR].OrderNr;