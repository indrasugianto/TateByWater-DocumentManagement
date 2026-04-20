SELECT qry_invoices_summary.CaseID, Last(tbl_InvoiceSent.InvSent) AS LastOfInvSent
FROM qry_invoices_summary INNER JOIN tbl_InvoiceSent ON qry_invoices_summary.CaseID = tbl_InvoiceSent.CaseID
GROUP BY qry_invoices_summary.CaseID, tbl_InvoiceSent.[TK Sent]
HAVING (((tbl_InvoiceSent.[TK Sent])=No));