SELECT qryInvoiceAttachRPT.CaseID, qryInvoiceAttachRPT.Bill_ID, qryInvoiceAttachRPT.IANumber, qryInvoiceAttachRPT.Tdate, qryInvoiceAttachRPT.Description, qryInvoiceAttachRPT.Tatty, qryInvoiceAttachRPT.Rate, qryInvoiceAttachRPT.Time_, Nz([time_],0)*Nz([rate],0) AS Amount
FROM qryInvoiceAttachRPT
ORDER BY qryInvoiceAttachRPT.Tdate;