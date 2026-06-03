SELECT qry_current_invoice.CaseID, qry_current_invoice.OrderNr, qry_current_invoice.Payment, qry_current_invoice.Pay_Outlay, qry_current_invoice.Date2
FROM qry_current_invoice
WHERE (((qry_current_invoice.Payment)<>0));