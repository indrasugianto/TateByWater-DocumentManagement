SELECT p.Name, p.FileNumber, p.MatterID, p.CaseID, p.Date2, p.Pay_Outlay, p.Charge, p.Payment, p.FirmPrepaid, p.InsertPymt, p.AdvancedLegal, p.SSMA_TimeStamp, p.Orig_Atty, p.Case_Letter, p.CodeVal, p.Creimb, CCur(Nz (t.SumOfBalance_agg, 0)) AS SumOfBalance
FROM vw_advanced_payments AS p LEFT JOIN (SELECT
            CaseID,
            Max(SumOfBalance) AS SumOfBalance_agg
        FROM
            qryTrustAccountBalanceTotals
        GROUP BY
            CaseID
    )  AS t ON p.CaseID = t.CaseID
ORDER BY p.Date2 DESC;