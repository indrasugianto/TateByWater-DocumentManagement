SELECT qryAdvLegalFees.CaseID, Sum(qryAdvLegalFees.Charge) AS SumOfCharge, qryAdvLegalFees.AdvancedLegal
FROM qryAdvLegalFees
GROUP BY qryAdvLegalFees.CaseID, qryAdvLegalFees.AdvancedLegal;