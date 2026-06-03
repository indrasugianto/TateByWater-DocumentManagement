SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName
FROM tblDropD
WHERE (((tblDropD.FieldName)="CompltMethod"))
ORDER BY tblDropD.SortOrder;