' Component: ModUpload
' Type: module
' Lines: 101
' ============================================================

'---------------------------------------------------------------------------------------
' Module    : modUpload
' Author    : EAC
' Date      : 7/24/2013
' Purpose   : This module is created for uploading a file.
'---------------------------------------------------------------------------------------

Option Compare Database

Public Enum ImportType
    Excel
    Text
End Enum

''******************************INSTRUCTIONS ON HOW TO USE******************************
'' Step 1. Prepare a table where we have to import the final data into
'''''''''''Main table should have a column UploadID with a number Data Type
'''''''''''Main table should Also have a column UploadedAt with a Date Data Type
'''''''''''Main table should Also have a column FileName with a Text Data Type
'' Step 2. Prepare a temporary table, with expected columns and text data type for all the columns in this table
'' Step 3. Prepare a query, which would filter the temporary data. Like extra rows, and the headings etc.
'' Step 4. Prepare a query, which would insert the data from temporary table to the main table.
'' Step 5. Call the Upload function to upload
''*************************************************************************************


'---------------------------------------------------------------------------------------
' Procedure : Upload
' Author    : EAC
' Date      : 7/24/2013
' Purpose   : This function is the only entry point for uploading a file into Db
'---------------------------------------------------------------------------------------
'

Public Function Upload(FileName As String, _
                  TemporaryTableName As String, _
                  MainTableName As String, _
                  QueryToInsert As String, _
                  Optional QueryToFilter As String = "", _
                  Optional DeleteTempData As Boolean = False, _
                  Optional impType As ImportType = ImportType.Excel) As Boolean
    
On Error GoTo Upload_Error


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Logic:
'   Make sure you have a table ready for importing data in a temp table.
'  -- This temp table should have exactly equal columns as in the source file
'  -- but with a text data type.
'
'   Make sure you have the query to Insert and to Filter ready
'
'   Import the table in to temp
'
'   Clear the junk datta
'
'   Insert the table from
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

   On Error GoTo 0
   
   'transfer the spreadsheet
   If impType = Text Then
       DoCmd.TransferText acLinkDelim, , TemporaryTableName, FileName, False
    Else
       DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, TemporaryTableName, FileName, False
   End If
   'file transferred. Filter the data
   If QueryToFilter <> "" Then CurrentDb.Execute QueryToFilter ' you only filter the data when you need to. If it is blank, it means that we dont hve to filter the data. And straight away fix it
   
   
   Dim ID As Integer
   
   ID = Nz(DMax("UploadID", MainTableName), 0) + 1
   
   
   'now insert the query
   'this query is a required field, if there is any error, then it has to be fixed before we go ahead
   CurrentDb.Execute QueryToInsert
   
   CurrentDb.Execute "update [" & MainTableName & "] set uploadID = " & ID & " where UploadID is null"
   CurrentDb.Execute "update [" & MainTableName & "] set uploadedAt = #" & Format(Now, "mm-dd-yyyy") & "# where UploadID = " & ID
   CurrentDb.Execute "update [" & MainTableName & "] set FileName = '" & FileName & "' where UploadID = " & ID
   
   'delete temporary data if it is required
   
   If DeleteTempData Then CurrentDb.Execute "delete * from " & TemporaryTableName ' it will clear the temporary table
   
   Upload = True
   
   Exit Function

Upload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Upload of Module modUpload"
End Function

