SELECT tblTimeTableDetail.Time_ID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Time_, tblTimeTableDetail.Rate, Nz([Time_],0)*Nz([rate],0) AS Amount
FROM tblTimeTableDetail;