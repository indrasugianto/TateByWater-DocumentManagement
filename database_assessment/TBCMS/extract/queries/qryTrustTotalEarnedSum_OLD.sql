SELECT Sum(qryTrustTotalEarned.SumOfCredit) AS SumOfSumOfCredit, tblCase.CaseID
FROM qryTrustTotalEarned INNER JOIN tblCase ON qryTrustTotalEarned.CaseID = tblCase.CaseID
GROUP BY tblCase.CaseID;