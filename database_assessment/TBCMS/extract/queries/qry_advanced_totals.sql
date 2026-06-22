SELECT tblCase.CaseID, [Matter and AR].Charge, [Matter and AR].FirmPrepaid
FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID
WHERE ((([Matter and AR].FirmPrepaid)=Yes));