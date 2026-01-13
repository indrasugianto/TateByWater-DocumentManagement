' Component: mod_create_table_with_all_db_schema
' Type: module
' Lines: 112
' ============================================================

Option Compare Database

Sub GetField2Description()
    '**********************************************************
    'Purpose:   1) Deletes and recreates a table (tblFields)
    '           2) Queries table MSysObjects to return names of
    '              all tables in the database
    '           3) Populates tblFields
    'Coded by:  raskew
    'Inputs:    From debug window:
    '           Call GetField2Description
    'Output:    See tblFields
    '**********************************************************
    
    Dim db As Database, td As TableDef
    Dim rs As Recordset, rs2 As Recordset
    Dim Test As String, NameHold As String
    Dim typehold As String, SizeHold As String
    Dim fielddescription As String, tName As String
    Dim n As Long, i As Long
    Dim fld As Field, strSQL As String
    n = 0
    Set db = CurrentDb
    ' Trap for any errors.
        On Error Resume Next
    tName = "tblFields"
    
    'Does table "tblFields" exist?  If true, delete it;
    DoCmd.SetWarnings False
       DoCmd.DeleteObject acTable, "tblFields"
    DoCmd.SetWarnings True
    'End If
    'Create new tblTable
    db.Execute "CREATE TABLE tblFields(Object TEXT (55), FieldName TEXT (55), FieldType TEXT (20), FieldSize Long, FieldAttributes Long, FldDescription TEXT (20));"
    
    strSQL = "SELECT MSysObjects.Name, MSysObjects.Type From MsysObjects WHERE"
    strSQL = strSQL + "((MSysObjects.Type)=1)"
    strSQL = strSQL + "ORDER BY MSysObjects.Name;"
    
    Set rs = db.OpenRecordset(strSQL)
    If Not rs.BOF Then
       ' Get number of records in recordset
       rs.MoveLast
       n = rs.RecordCount
       rs.MoveFirst
    End If
    
    Set rs2 = db.OpenRecordset("tblFields")
    
    For i = 0 To n - 1
      fielddescription = " "
      Set td = db.TableDefs(i)
        'Skip over any MSys objects
        If Left(rs!Name, 4) <> "MSys" And Left(rs!Name, 1) <> "~" Then
           NameHold = rs!Name
           On Error Resume Next
           For Each fld In td.Fields
              fielddescription = fld.Name
              typehold = FieldType(fld.Type)
              SizeHold = fld.Size
              rs2.AddNew
              rs2!Object = NameHold
              rs2!FieldName = fielddescription
              rs2!FieldType = typehold
              rs2!FieldSize = SizeHold
              rs2!FieldAttributes = fld.Attributes
              rs2!FldDescription = fld.Properties("description")
              rs2.Update
           Next fld
    
           Resume Next
        End If
        rs.MoveNext
    Next i
    rs.Close
    rs2.Close
    db.Close
End Sub

Function FieldType(intType As Integer) As String
    
    Select Case intType
        Case dbBoolean
            FieldType = "dbBoolean"    '1
        Case dbByte
            FieldType = "dbByte"       '2
        Case dbInteger
            FieldType = "dbInteger"    '3
        Case dbLong
            FieldType = "dbLong"       '4
        Case dbCurrency
            FieldType = "dbCurrency"   '5
        Case dbSingle
            FieldType = "dbSingle"     '6
        Case dbDouble
            FieldType = "dbDouble"     '7
        Case dbDate
            FieldType = "dbDate"       '8
        Case dbBinary
            FieldType = "dbBinary"     '9
        Case dbText
            FieldType = "dbText"       '10
        Case dbLongBinary
            FieldType = "dbLongBinary" '11
        Case dbMemo
            FieldType = "dbMemo"       '12
        Case dbGUID
            FieldType = "dbGUID"       '15
    End Select

End Function
