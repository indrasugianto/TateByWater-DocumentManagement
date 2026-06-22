SELECT tblTimeTableDetail.Bill_ID, Sum(tblTimeTableDetail.Time_) AS SumOfTime_, tblTimeTableDetail.Tatty
FROM tblTimeTableDetail
GROUP BY tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Tatty;