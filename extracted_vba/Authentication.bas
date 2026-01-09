' Component: Authentication
' Type: module
' Lines: 52
' ============================================================

Option Compare Database

Public LoggedUser As User
Public ACSType As AccessType
Public Const StartFormName As String = "frmLogin"
Public Const CheckAuthentication As Boolean = True

Public Function IsLoggedIn() As Boolean
    If LoggedUser Is Nothing Then
        IsLoggedIn = False
    Else
        IsLoggedIn = True
    End If
End Function

Public Sub LogOut()
    Set LoggedUser = Nothing
End Sub

Public Function Login(uid As String, PWD As String, Optional showMsg As Boolean = True) As Boolean
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT tblUsers.*, tblAccessType.AccessType, tblAccessType.AccessDescription, tblAccessType.AdminPane FROM tblAccessType " & _
            " INNER JOIN tblUsers ON tblAccessType.AccessType = tblUsers.Access where tblUsers.UserID = '" & uid & "' and tblUsers.PWD='" & PWD & "' ;"

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    
    If rs.EOF = False Then
        BindObjects rs
        Login = True
    Else
        Login = False
        If showMsg Then
            ShowMessage "The user id and password don't match. Please re-enter the credentials and try again."
        End If
    End If
End Function

Public Sub BindObjects(record As Recordset)
    
    'Stop
    
    Set LoggedUser = New User
    LoggedUser.ID = record(0)
    LoggedUser.LoginID = record(1)
    LoggedUser.AccessType = record(3)
    
    Set ACSType = New AccessType
    ACSType.Description = record(5)
    ACSType.AdminPane = record(6)
End Sub