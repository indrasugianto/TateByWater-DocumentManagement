' Component: FormUtils
' Type: module
' Lines: 273
' ============================================================

Option Compare Database


'''Logic for Authentication

Public Function Form_Load_New(frmObject As Form, Optional AuthenticationRequired As Boolean) As Boolean
 Dim frm As Form
    
   On Error GoTo Form_Load_New_Error

    'assign title, if found in the titles tables.

    Dim Title As String
    Dim formName As String
    
    formName = frmObject.Name
    
    Title = GetTitle(formName)
    
    If AuthenticationRequired = True And Authentication.CheckAuthentication = True Then
        Dim MinimumAccessRequired_Show As Integer
        Dim MinimumAccessRequired_Edit As Integer
        
        
        On Error Resume Next
        
        MinimumAccessRequired_Show = DLookup("MinimumAccess_Show", "tblFormAccessMapping", "FormName='" & formName & "'")
        MinimumAccessRequired_Edit = DLookup("MinimumAccess_Edit", "tblFormAccessMapping", "FormName='" & formName & "'")
        
        On Error GoTo Form_Load_New_Error ' resetting error
        
        If LoggedUser Is Nothing Then
            Form_Load_New = False
            'try to open the login window
            DoCmd.openform "frmLogin", , , , , acDialog
            
            'if the user has opted to Cancel, then we have no choice other than killing the application.
            
            If LoggedUser Is Nothing Then
                ShowMessage "Exiting"
                Application.Quit
            Else
                'let the execution flow like normal
            End If
            
        End If
        
        If LoggedUser.AccessType >= MinimumAccessRequired_Show Then
            'allowed
            
            'check if the edit access is given or not
            
            If LoggedUser.AccessType < MinimumAccessRequired_Show Then
                'hide
                frmObject.AllowEdits = False
                'now disable the necessary controls
                DisableControlsThroughTag frmObject
            End If
            
        Else
            'it means that the user doesn't have the privelleges to even view the window.
            Form_Load_New = False
            ShowMessage "You don't have sufficient privelleges to access this window."
            Util.CloseForm
            Exit Function
        End If
    
    End If

    If Title <> "" Then frmObject.Caption = Title

    If callCodeBehind Then
        With CodeContextObject
            If .CodeBehind_Form_Load = False Then
                Form_Load_New = False
            End If
        End With
    End If
    Form_Load_New = True
   On Error GoTo 0
   Exit Function

Form_Load_New_Error:
    Form_Load_New = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Module FormUtils"
End Function

Public Function Form_Load(Optional callCodeBehind As Boolean = False, Optional AuthenticationRequired As Boolean) As Boolean
    Dim frm As Form
    
   On Error GoTo Form_Load_Error

    'assign title, if found in the titles tables.

    Dim Title As String
    
    Title = GetTitle
    
    DoCmd.NavigateTo "acnavigationcategoryobjectType"
    DoCmd.RunCommand acCmdWindowHide
    
    If AuthenticationRequired = True Then
        Dim MinimumAccessRequired_Show As Integer
        Dim MinimumAccessRequired_Edit As Integer
        
        Dim formName As String
        
        formName = CodeContextObject.Name
        On Error Resume Next
        
        MinimumAccessRequired_Show = DLookup("MinimumAccess_Show", "tblFormAccessMapping", "FormName='" & formName & "'")
        MinimumAccessRequired_Edit = DLookup("MinimumAccess_Edit", "tblFormAccessMapping", "FormName='" & formName & "'")
        
        On Error GoTo Form_Load_Error ' resetting error
        
        If LoggedUser Is Nothing Then
            Form_Load = False
            Dim frmName As String
            frmName = CodeContextObject.Name
            'try to open the login window
            DoCmd.openform "frmLogin", , , , , acDialog
            
            'if the user has opted to Cancel, then we have no choice other than killing the application.
            
            If LoggedUser Is Nothing Then
                ShowMessage "Exiting"
                Application.Quit
            Else
                'let the execution flow like normal
            End If
            
        End If
        
        If LoggedUser.AccessType >= MinimumAccessRequired_Show Then
            'allowed
            
            'check if the edit access is given or not
            
            If LoggedUser.AccessType < MinimumAccessRequired_Edit Then
                'hide
                CodeContextObject.AllowEdits = False
                'now disable the necessary controls
                DisableControlsThroughTag CodeContextObject
            End If
            
        Else
            'it means that the user doesn't have the privelleges to even view the window.
            Form_Load = False
            ShowMessage "You don't have sufficient privelleges to access this window."
            Util.CloseForm
            Exit Function
        End If
        
    End If

    If Title <> "" Then CodeContextObject.Caption = Title

    If callCodeBehind Then
        With CodeContextObject
            If .CodeBehind_Form_Load = False Then
                Form_Load = False
            End If
        End With
    End If

   On Error GoTo 0
   Exit Function

Form_Load_Error:
    Form_Load = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Module FormUtils"
End Function

''''if you want to disable few buttons like Submit/Edit etc. if user doesn;t have edit access,
''''then tag the control by setting Tag = "Disable"


Public Sub DisableControlsThroughTag(fmrObject As Object)
    
    Dim control_ As Control
    
    For Each control_ In fmrObject.Controls
        On Error Resume Next
        If control_.Tag = "Disable" Then
            control_.Enabled = False
        End If
    Next control_
    
End Sub

Public Function GetTitle(Optional Name As String = "") As String
    
    
    If Name = "" Then Name = CodeContextObject.Name
    
    On Error Resume Next
    
    GetTitle = Nz(DLookup("Title", "tblTitles", "FormName='" & Name & "'"), "")
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : ValidateCombo
' Author    : EAC
' Date      : 5/21/2013
' Purpose   : Check : index is selected or not
'---------------------------------------------------------------------------------------
'
Public Function ValidateCombo(control_ As ComboBox, Optional showMsg As Boolean = True) As Boolean

    ValidateCombo = True
    Dim controlName As String

    If IsNull(control_) Then
        ValidateCombo = False
        If showMsg Then
            controlName = control_.Tag
            ShowMessage "Please enter a value in " & controlName & " and try again."
        End If
    Else
        If control_.ListIndex < 0 Then
            ValidateCombo = False
            'not selected
            If showMsg Then
                controlName = control_.Tag
                ShowMessage "Please enter a value in " & controlName & " and try again."
            End If
        End If
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ValidateTB
' Author    : EAC
' Date      : 5/21/2013
' Purpose   : Checks if there is a value in textbox or not
'---------------------------------------------------------------------------------------
'
Public Function ValidateTB(control_ As TextBox, Optional showMsg As Boolean = True) As Boolean

    ValidateTB = True

    If IsNull(control_) Then
        'not selected
        ValidateTB = False
        If showMsg Then
            Dim controlName As String
            controlName = control_.Tag
            ShowMessage "Please enter a value in " & controlName & " and try again."
        End If
    End If
End Function


Public Function OnclickEdit(strfrmName As String, eid As Variant)
    Dim stDocName As String
    Dim stlinkcriteria As String
    
    stDocName = strfrmName

    If IsNumeric(eid) Then
        stlinkcriteria = "[ID]=" & Eval(eid)
    Else
        stlinkcriteria = "[ID]='" & (eid) & "'"
    End If
    
    DoCmd.openform stDocName, , , stlinkcriteria, acFormEdit, acDialog

End Function
