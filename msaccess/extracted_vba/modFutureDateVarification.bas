' Component: modFutureDateVarification
' Type: module
' Lines: 8
' ============================================================

Option Compare Database

Public Function DateVarifier(ByRef dt As Date) As Boolean
    If dt > Date Then
        MsgBox "Future dates can't be accepted.", vbInformation, "TB CMS: Invalid Input"
        DateVarifier = True
    End If
End Function