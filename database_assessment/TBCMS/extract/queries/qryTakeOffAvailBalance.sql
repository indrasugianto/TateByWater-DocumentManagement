SELECT CaseID, [Balance]-Nz([SumOfUnclearedDeposits],0) AS AvailBalance
FROM qryTakeOff_A;