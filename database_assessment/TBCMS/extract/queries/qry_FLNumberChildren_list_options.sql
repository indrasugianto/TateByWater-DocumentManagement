SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName
FROM tblDropD
WHERE (((tblDropD.FieldName)="NumberChildren"))
ORDER BY tblDropD.SortOrder;