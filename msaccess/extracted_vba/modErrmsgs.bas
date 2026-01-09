' Component: modErrmsgs
' Type: module
' Lines: 5
' ============================================================

Option Compare Database

Public Function Mbox(strFldName As String, lngId As Long)
    MsgBox "'" & strFldName & "' " & DLookup("Message", "errMsgs", "ID=" & lngId) & "'", vbInformation, DLookup("Title", "errMsgs", "ID=" & lngId)
End Function