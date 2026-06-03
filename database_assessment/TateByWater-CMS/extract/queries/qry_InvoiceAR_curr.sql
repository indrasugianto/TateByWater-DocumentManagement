SELECT qry_current_invoice.Date2, qry_current_invoice.Pay_Outlay, qry_current_invoice.Charge, qry_current_invoice.OrderNr, qry_current_invoice.CaseID
FROM qry_current_invoice
WHERE (((qry_current_invoice.Charge)>0));