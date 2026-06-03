SELECT qryInvoiceRPT1.CaseID, qryInvoiceRPT1.MatterID, qryInvoiceRPT1.Charge, qryInvoiceRPT1.Payment, qryInvoiceRPT1.CaseOpenDate, qryInvoiceRPT1.Date2, fncRunningDebit([CaseID],[Date2],[MatterID]) AS RunningDebit, fncRunningCredit([CaseID],[Date2],[MatterID]) AS RunningCredit, [RunningDebit]-[RunningCredit] AS RunningBalance
FROM qryInvoiceRPT1
WHERE (((qryInvoiceRPT1.CaseID)=11));