SELECT tblTakeOffMonth.TakeOffDate
FROM tblTakeOffMonth INNER JOIN tblTakeOff ON tblTakeOffMonth.TakeOffMonthID = tblTakeOff.TakeOffMonthID;