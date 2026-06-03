SELECT qryARCredits.CaseID, Sum(qryARCredits.Payment) AS SumOfPayment
FROM qryARCredits
GROUP BY qryARCredits.CaseID;