SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName
FROM tblDropD
WHERE (((tblDropD.FieldName)="ChildCustodian"))
ORDER BY tblDropD.SortOrder;