SELECT qryInvoiceRPT1.*, [Runningbalance]+[Retainer] AS RetBal
FROM qryInvoiceRPT1
ORDER BY qryInvoiceRPT1.MatterID DESC;