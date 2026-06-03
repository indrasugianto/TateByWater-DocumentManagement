SELECT tblTimeTableDetail.Bill_ID, Sum(tblTimeTableDetail.Time_) AS SumOfTime_
FROM tblTimeTableDetail
GROUP BY tblTimeTableDetail.Bill_ID;