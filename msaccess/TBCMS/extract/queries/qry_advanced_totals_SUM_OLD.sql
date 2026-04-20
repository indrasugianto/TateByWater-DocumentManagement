SELECT qry_advanced_totals.CaseID, Sum(qry_advanced_totals.Charge) AS SumOfCharge, qry_advanced_totals.FirmPrepaid
FROM qry_advanced_totals
GROUP BY qry_advanced_totals.CaseID, qry_advanced_totals.FirmPrepaid;