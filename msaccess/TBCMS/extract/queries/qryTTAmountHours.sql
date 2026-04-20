SELECT tblTimeTableDetail.Time_ID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Tatty, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, Sum(Nz([Time_],0)) AS Amount
FROM tblTimeTableDetail;