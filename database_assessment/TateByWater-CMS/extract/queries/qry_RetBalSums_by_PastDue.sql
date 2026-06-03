SELECT qryOutstandingARRPT1.[Past Due], qryOutstandingARRPT1.Orig_Atty, Sum(qryOutstandingARRPT1.RetBal) AS SumOfRetBal
FROM qryOutstandingARRPT1
GROUP BY qryOutstandingARRPT1.[Past Due], qryOutstandingARRPT1.Orig_Atty;