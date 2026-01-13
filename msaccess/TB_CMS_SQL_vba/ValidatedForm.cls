' Component: ValidatedForm
' Type: class
' Lines: 10
' ============================================================

Option Compare Database

Public WithEvents MyForm As Access.Form

Private Sub MyForm_Error(DataErr As Integer, Response As Integer)
    If DataErr = 2113 Then
        ShowMessage "Please enter proper data for " & Screen.ActiveControl.Name
        Response = 0
    End If
End Sub