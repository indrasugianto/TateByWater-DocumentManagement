SELECT tblTakeOffMonth.TakeOffMonthID, tblTakeOffMonth.TakeOffDate, Year([TakeOffDate]) AS YearOnly, Month([TakeOffDate]) AS MonthOnly
FROM tblTakeOffMonth;