SELECT tblDropD.Code, tblDropD.CodeVal, tblDropD.FieldName, tblDropD.SortOrder
FROM tblDropD
WHERE (((tblDropD.FieldName)="Orig_Atty"))
ORDER BY tblDropD.SortOrder;