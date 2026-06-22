SELECT tblCase.CaseID, Sum([Matter and AR].Charge) AS SumOfCharge
FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID
GROUP BY tblCase.CaseID, [Matter and AR].FirmPrepaid
HAVING ((([Matter and AR].FirmPrepaid)=Yes));