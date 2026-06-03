SELECT vwTrustAccount.CaseID, vwTrustAccount.SumOfDebit, vwTrustAccount.SumOfCredit, vwTrustAccount.Balance, vwTrustAccount.OrderNr, vwTrustAccount.TrustAccountID
FROM vwTrustAccount
ORDER BY vwTrustAccount.CaseID, vwTrustAccount.OrderNr;