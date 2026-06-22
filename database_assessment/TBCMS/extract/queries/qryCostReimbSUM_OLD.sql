SELECT qryCostReimb.CaseID, Sum(qryCostReimb.Credit) AS SumOfCredit
FROM qryCostReimb
GROUP BY qryCostReimb.CaseID;