SELECT [TB Time Keeping].CaseID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, [Rate]*[Time_] AS Total, [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber
FROM [TB Time Keeping] INNER JOIN tblTimeTableDetail ON [TB Time Keeping].Bill_ID = tblTimeTableDetail.Bill_ID
WHERE ((([TB Time Keeping].[Bill Closed])=No));