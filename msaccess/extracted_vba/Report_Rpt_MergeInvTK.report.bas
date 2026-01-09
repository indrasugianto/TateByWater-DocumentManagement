' Component: Report_Rpt_MergeInvTK
' Type: document
' Lines: 38
' ============================================================

Option Compare Database

Private Sub Charge19_Click()

End Sub

'Private Sub Report_Load()
'
'    If Me.Dirty Then Me.Dirty = False
'    Call reorderByDate
'
'End Sub
'
'Sub reorderByDate()
'    Dim rs As Recordset
'
'    strSQL = "select * from [Matter and AR] where CaseID = " & Report_Invoice2.CaseID & " order by Date2"
'    Debug.Print strSQL
'
'    Set rs = CurrentDb.OpenRecordset(strSQL)
'    Dim i As Integer
'    i = 0
'    Do Until rs.EOF
'        i = i + 1
'
'        strSQL = "Update [Matter and AR] set OrderNr = " & i & " where MatterID =" & rs("MatterID")
'        Debug.Print "Date is: " & rs("date2")
'        Debug.Print strSQL
'        CurrentDb.Execute strSQL
'
'        rs.MoveNext
'    Loop
'
'End Sub

Private Sub Report_NoData(Cancel As Integer)
Cancel = True
End Sub