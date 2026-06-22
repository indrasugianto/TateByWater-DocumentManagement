SELECT vwMatter.MatterID, vwMatter.Date2, vwMatter.CaseID, vwMatter.SumOfCharge, vwMatter.SumOfPayment, vwMatter.Balance, vwMatter.OrderNr
FROM vwMatter
ORDER BY vwMatter.CaseID, vwMatter.OrderNr;