SELECT qry_time_table_totals_atty.Tatty, Sum(qry_time_table_totals_atty.SumOfAmount) AS SumOfSumOfAmount, [TB Time Keeping].CaseID
FROM qry_time_table_totals_atty INNER JOIN [TB Time Keeping] ON qry_time_table_totals_atty.Bill_ID = [TB Time Keeping].Bill_ID
GROUP BY qry_time_table_totals_atty.Tatty, [TB Time Keeping].CaseID;