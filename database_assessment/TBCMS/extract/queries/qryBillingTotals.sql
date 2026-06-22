SELECT tblTimeTableDetail.Bill_ID, Sum(Nz([Time_],0)*Nz([Rate],0)) AS Amount
FROM tblTimeTableDetail
GROUP BY tblTimeTableDetail.Bill_ID;