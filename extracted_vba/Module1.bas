' Component: Module1
' Type: module
' Lines: 44
' ============================================================

Option Compare Database

Public dtST1            As Date
Public dtST2            As Date
Public lngBill_ID       As Long
Public lngCaseID        As Long
Public lngChnc          As Long
Public StatementofTrustAccount_Filter As Boolean

Public Function getSTDT() As Date
    getSTDT = dtST1
End Function
Public Function getENDT() As Date
    getENDT = dtST2
End Function

Public Function GetCaseID() As Long
    GetCaseID = lngCaseID
End Function

Public Function GetRetainer(CaseID As Long)
    GetRetainer = Nz(DLookup("retainer", "tblCase", "CaseID=" & GetCaseID()), 0)
End Function

Public Function GetBill_ID() As Long
    GetBill_ID = lngBill_ID
End Function

Public Sub ErrMsg(ProcName As String)
    MsgBox "An error has occurred in the application at Proc/Func- '" & ProcName & "'." _
         & vbNewLine & vbNewLine & "Error Number:-     " & Err.Number & vbNewLine & "Error Description:-" & Err.Description, vbCritical, "Error Message"
End Sub

Public Sub OpenFormHomeScreen(frmName As String)
On Error GoTo ErrHnd
    DoCmd.openform frmName, acNormal, , , acFormEdit, acWindowNormal
Exit Sub
ErrHnd:
    ErrMsg "OpenFormHomeScreen"
End Sub



 