' Component: Relinking
' Type: class
' Lines: 131
' ============================================================

Option Compare Database

Dim db As Object
Const dbName As String = "Login_BE1.accdb"

Private Sub Class_Initialize()
    Set db = CurrentDb
End Sub



'---------------------------------------------------------------------------------------
' Procedure : ReLink
' Author    : EAC
' Date      : 5/21/2013
' Purpose   : Relinks the DB
'---------------------------------------------------------------------------------------
'
Public Function ReLink(Optional ShowMessage As Boolean = True) As Boolean
    
    Dim bEnd_dbName As String
    
   On Error GoTo ReLink_Error

    bEnd_dbName = GetLinkName
    
    'if the link exists, try to find it it exists or not.
    
    If PathExists(bEnd_dbName) Then
        'fine, stop the processing
        ReLink = True
    Else
        'the path doesn't exist. Try to find the link in same folder
        
        Dim mainPath As String
        
        
        mainPath = CurrentProject.path & "\" & dbName
        
        
        If PathExists(mainPath) Then
            'link it
            LinkTables mainPath
            
            If (ShowMessage) Then
                Util.ShowMessage "The tables are relinked to " & mainPath & "."
            End If
            
             ReLink = True
            
        Else
            'this is serious. We cannot find the DB. Get it browsed
            'confirm from Client.
            
            DoCmd.openform "frmBrowse_backEnd", acNormal, , , , acDialog
            
            If Context.YesNo_Value = 1 Then
                'clicked OK
                LinkTables SelectedPath
                If (ShowMessage) Then
                    Util.ShowMessage "The tables are relinked to " & SelectedPath & "."
                End If
                ReLink = True

            Else
                ReLink = False
            End If
            
            
        End If
        
    End If

   On Error GoTo 0
   Exit Function

ReLink_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReLink of Class Module Relinking"
    ReLink = True ' means stop the function
End Function

Private Function PathExists(path As String) As Boolean
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject

    PathExists = FSO.FileExists(path)

    Set FSO = Nothing
End Function

Public Function GetLinkName() As String
    Dim tblDef As TableDef
    For Each tblDef In db.TableDefs
        If tblDef.Connect <> "" Then
            GetLinkName = ExtractDBName(tblDef.Connect)
            Exit Function
        End If
    Next tblDef
End Function

Public Function LinkTables(newDBName As String) As String
    Dim tblDef As TableDef
    For Each tblDef In db.TableDefs
        If tblDef.Connect <> "" Then
            Dim newConnect As String
            newConnect = ExtractDBName(tblDef.Connect)
            tblDef.Connect = Replace(tblDef.Connect, newConnect, newDBName)
            tblDef.RefreshLink
        End If
    Next tblDef
End Function

Public Function ExtractDBName(stringConnect As String) As String
    Dim varArray() As String
    Dim value
    
    varArray = Split(stringConnect, ";")
    
    For Each value In varArray
        If InStr(1, value, "Database=") > 0 Then
            ExtractDBName = Replace(value, "Database=", "")
        End If
    Next value
End Function

Public Function IsLinked() As Boolean
    On Error Resume Next
    IsLinked = Len(GetLinkName) > 0
End Function
