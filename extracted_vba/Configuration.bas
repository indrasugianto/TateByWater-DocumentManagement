' Component: Configuration
' Type: module
' Lines: 75
' ============================================================

Option Compare Database

'////////////////////////////////////////////////////////////////////////
'/Here we will be storing the configurations, like for how many users  //
'/we will be showing the Pane etc                                      //
'/                                                                     //
'/                                                                     //
'////////////////////////////////////////////////////////////////////////

Function ap_DisableShift()
'This function disable the shift at startup. This action causes
'the Autoexec macro and Startup properties to always be executed.

On Error GoTo errDisableShift

Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

'This next line disables the shift key on startup.
db.Properties("AllowByPassKey") = False

'The function is successful.
Exit Function

errDisableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", _
dbBoolean, False)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function

Function ap_EnableShift()
'This function enables the SHIFT key at startup. This action causes
'the Autoexec macro and the Startup properties to be bypassed
'if the user holds down the SHIFT key when the user opens the database.

On Error GoTo errEnableShift

Dim db As DAO.Database
Dim prop As DAO.Property
Const conPropNotFound = 3270

Set db = CurrentDb()

'This next line of code disables the SHIFT key on startup.
db.Properties("AllowByPassKey") = True

'function successful
Exit Function

errEnableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", _
dbBoolean, True)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function