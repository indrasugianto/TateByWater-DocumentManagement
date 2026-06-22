SELECT tbl_InvoiceSent.CaseID, Last(tbl_InvoiceSent.InvSent) AS LastOfInvSent
FROM tbl_InvoiceSent
GROUP BY tbl_InvoiceSent.CaseID;