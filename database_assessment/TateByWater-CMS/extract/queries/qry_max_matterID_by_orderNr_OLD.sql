SELECT Max([Matter and AR].OrderNr) AS MaxOfOrderNr, Max([Matter and AR].MatterID) AS MaxOfMatterID, [Matter and AR].CaseID
FROM [Matter and AR]
GROUP BY [Matter and AR].CaseID;