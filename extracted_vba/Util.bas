' Component: Util
' Type: module
' Lines: 210
' ============================================================

Option Compare Database

'////////////////////////////////////////////////////////////////////////
'/There has to be an entry point, where you will be writing all the    //
'/codes like relinking etc, Main function should be the entry point    //
'/                                                                     //
'/                                                                     //
'////////////////////////////////////////////////////////////////////////

Public Function Main() ' entry point of application

    Dim linking As Relinking

    Set linking = New Relinking
    
    If linking.ReLink(True) = False Then
        message = "Couldn't find the Backend DB. Please locate and link DB Manually."
        ShowMessage message
    End If
    
    Call DisableProperties
    If StartFormName <> "" Then
        DoCmd.openform StartFormName, acNormal
    End If
    
    If Debugging = True Then
        
    End If
    
End Function

Public Sub ShowMessage(messageText As String, Optional messageText1 As String = "")
    message = messageText
    Message1 = messageText1
    DoCmd.openform "frmOkAlert", acNormal, , , , acDialog
End Sub

Public Sub ShowYesNo(messageText As String, Optional messageText1 As String = "")
    message = messageText
    Message1 = messageText1
    DoCmd.openform "frmYesNoAlert", acNormal, , , , acDialog
End Sub

Public Function GetMessageText(acMsgType As MsgType)
    Dim strText As String
    strText = Nz(DLookup("MessageName", "Tblmsgbox", "ID=" & acMsgType), "")
    GetMessageText = strText
End Function

Public Function BrowseDB(Optional Title As String = "Select Access DB", Optional Description As String = "Access Database", _
                            Optional Filter As String = "*.accdb") As String

    Dim d As FileDialog

    Set d = Application.FileDialog(msoFileDialogFilePicker)
    
    d.AllowMultiSelect = False
    d.Filters.Clear
    d.Title = Title
    d.Filters.Add Description, Filter
    d.show
    On Error Resume Next
    BrowseDB = d.SelectedItems(1)
    
End Function

Public Function openform(Optional frmName As String = "")
        If frmName = "" Then
        ' do nothing
        Else
            DoCmd.openform frmName, acNormal
        End If
End Function
Public Function CloseForm(Optional frmName As String = "")

    If frmName = "" Then frmName = CodeContextObject.Name
    DoCmd.Close acForm, frmName, acSaveNo
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetColumnHeading
' Author    : EAC
' Date      : 5/22/2013
' Purpose   : Returns the column heading for Excel Column, like A for 1 , 27 for AA
'---------------------------------------------------------------------------------------
'
Public Function GetColumnHeading(colNum As Integer) As String
    On Error Resume Next
    Dim max As Integer
    
    max = UBound(ColumnHeadings)
    
    On Error GoTo 0
    
    If max = 0 Then
        'load array
        ColumnHeadings = Split(ColumnNames, ",")
    End If
    
    GetColumnHeading = Trim(ColumnHeadings(colNum - 1))
    
End Function

Public Sub DisableAll()
    LogOut
End Sub

Public Sub showNavigationPane(show As Boolean)
    'To-do: If show is true then the navigation should be visible
            ' if false, then it should not be visible.
            
    On Error GoTo ErrHandler
            'Stop
    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    
    If show Then
        'DoCmd.RunCommand acCmdWindowUnhide
        Call DoCmd.SelectObject(acTable, , True)
    Else
        DoCmd.RunCommand acCmdWindowHide
    End If
    Exit Sub
ErrHandler:
    Debug.Print Err.Description
End Sub

Public Function SetProperties(PropName As String, PropType As Variant, PropValue As Variant) As Integer
    On Error GoTo Err_SetProperties
    Dim db As Database, prop As Property
    'Dim db As DAO.Database, prop As DAO.Property (use in the old version prior 2007)
    Set db = CurrentDb
    db.Properties(PropName) = PropValue
    SetProperties = True
    Set db = Nothing
Exit_SetProperties:
    Exit Function
Err_SetProperties:
    If Err = 3270 Then 'case of property not found
        Set prop = db.CreateProperty(PropName, PropType, PropValue)
        db.Properties.Append prop
        Resume Next
    Else
        SetProperties = False
        MsgBox "Runtime Error # " & Err.Number & vbCrLf & vbLf & Err.Description
        Resume Exit_SetProperties
    End If
End Function


Public Sub DisableProperties()
    'disable all properties
    
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
'    'Stop
'
    Call showNavigationPane(False)
'
    SetProperties "AppTitle", dbText, CompanyName
'
'    SetProperties "StartUpShowDBWindow", dbBoolean, False
'    SetProperties "StartUpShowStatusBar", dbBoolean, False
'
'    SetProperties "AllowFullMenus", dbBoolean, False
    SetProperties "AllowSpecialKeys", dbBoolean, False
'    SetProperties "AllowShortcutMenus", dbBoolean, False
'    SetProperties "AllowToolbarChanges", dbBoolean, False
'    SetProperties "AllowBuiltInToolbars", dbBoolean, False
'
    SetProperties "AllowBypassKey", dbBoolean, False
    SetProperties "AllowBreakIntoCode", dbBoolean, False
'
'    Application.SetOption "ShowWindowsInTaskbar", False
'    Application.SetOption "Themed Form Controls", True
'    Application.SetOption "Show Startup Dialog Box", False
End Sub

Public Function EnableProperties()
    'Set all properties listed below back to normal by setting value to True
    
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    
'    'Stop
'
'    On Error GoTo ErrorHandler:
'
    Call showNavigationPane(True)
'
'    SetProperties "StartUpShowDBWindow", dbBoolean, True
'    SetProperties "StartUpShowStatusBar", dbBoolean, True
'    SetProperties "AllowFullMenus", dbBoolean, True
    SetProperties "AllowSpecialKeys", dbBoolean, True
'    SetProperties "AllowShortcutMenus", dbBoolean, True
'    SetProperties "AllowToolbarChanges", dbBoolean, True
'    SetProperties "AllowBuiltInToolbars", dbBoolean, True
'
    SetProperties "AllowBypassKey", dbBoolean, True
    SetProperties "AllowBreakIntoCode", dbBoolean, True
'
'    Exit Function
'ErrorHandler:
'    MsgBox Err.Description
End Function

Function isACCDE() As Boolean
    isACCDE = False
    If Right(CurrentDb.Name, 5) = "accde" Then
        isACCDE = True
    End If
End Function