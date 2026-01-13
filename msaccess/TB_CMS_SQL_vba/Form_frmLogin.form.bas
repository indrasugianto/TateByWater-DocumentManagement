' Component: Form_frmLogin
' Type: document
' Lines: 62
' ============================================================

Option Compare Database

Private Sub cmdNo_Click()
    If MsgBox("Are you sure you want to exit the application?", vbYesNo + vbCritical, "TB CMS: Exit Application?") = vbYes Then Util.CloseForm
End Sub

Private Sub cmdYes_Click()
    
    If Not FormUtils.ValidateCombo(UserID) Then
        GoTo ExitSub ' dont process if there is no user ID
    End If
    
    If Not FormUtils.ValidateTB(PWD) Then
        GoTo ExitSub ' dont process if there is no pwd
    End If
    
    If Authentication.Login(UserID, PWD) Then
        'awesome logged in.
        'bind the user id now
        'close the form
'        If LoggedUser.AccessType = 5 Then
'            DoCmd.NavigateTo "acNavigationCategoryObjectType"
'            DoCmd.RunCommand acCmdWindowUnhide
'            DoCmd.LockNavigationPane False
'        Else
'            DoCmd.NavigateTo "acNavigationCategoryObjectType"
'            DoCmd.RunCommand acCmdWindowHide
'            DoCmd.LockNavigationPane False
'        End If
        
        Util.CloseForm
    End If
    
ExitSub:
End Sub

Public Function CodeBehind_Form_Load() As Boolean
    YesNo_Value = -1
    
    lblMessage.Caption = message
    lblMessage2.Caption = Message1
    
    CodeBehind_Form_Load = True
End Function

Private Sub cmdYes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdYes_Click
End Sub

Private Sub Form_Close()
    If LoggedUser Is Nothing Then
        Application.Quit
    Else
        If ACSType.AdminPane Then
            Call EnableProperties
        Else
            Call DisableProperties
        End If
        openform "frmHome"
    End If
End Sub
