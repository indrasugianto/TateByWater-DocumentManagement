SELECT vw_last_invoice_sent.CaseID, vw_last_invoice_sent.LastOfInvSent
FROM vw_last_invoice_sent
GROUP BY vw_last_invoice_sent.CaseID, vw_last_invoice_sent.LastOfInvSent;