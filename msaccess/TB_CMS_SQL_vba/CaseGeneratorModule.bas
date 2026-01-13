' Component: CaseGeneratorModule
' Type: module
' Lines: 46
' ============================================================

Option Compare Database

'GH: it's wrong to update all field values!
'This function will not be used anymore.

'Public Function CaseGenerator()
'    Dim rs2             As Recordset
'    Dim rsFiltered      As Recordset
'    Dim MinYear         As Long
'    Dim MaxYear         As Long
'    Dim maxval          As Long
'
'    CurrentDb.Execute "update tblcase set [Number_] = null "
'
'    Set rs2 = CurrentDb.OpenRecordset("select * from tblcase")
'
'    MinYear = DMin("yr", "tblCase")
'    MaxYear = DMax("yr", "tblcase")
'
'    For intval = MinYear To MaxYear
'
'        rs2.filter = "yr= '" & intval & "'"
'        Set rsFiltered = rs2.OpenRecordset
'
'        If Not rsFiltered.EOF Then
'            rsFiltered.MoveLast
'            rsFiltered.MoveFirst
'
'            maxval = Nz(DMax("Number_", "tblCase", "Yr='" & intval & "'"), 0)
'
'            Do Until rsFiltered.EOF
'                rsFiltered.Edit
'
'                rsFiltered!Number_ = maxval + 1
'                rsFiltered.Update
'                rsFiltered.MoveNext
'                maxval = maxval + 1
'            Loop
'        End If
'        rsFiltered.Close
'        Set rsFiltered = Nothing
'
'    Next
'    rs2.Close
'    Set rs2 = Nothing
'End Function