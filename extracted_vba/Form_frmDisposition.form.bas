' Component: Form_frmDisposition
' Type: document
' Lines: 22
' ============================================================

Option Compare Database

'Private Sub cmdInsertTotalEarned_Click()
'    If frmClientLedger.Closed.value = -1 Then
'    strSQL = "Update [tblDisposition] set [TotalEarned] = #" & txtTotalEarned & "# where CaseID=" & Me.CaseID
'    'strSQL = "Update [tblDisposition] set [TotalEarned] = #" & Format(Date, "yyyy-MM-dd") & "# where CaseID=" & Me.CaseID
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL
'    Me.Requery
'    Else
'    MsgBox "File Must be Closed to Insert Total Earned Fee.", , "TB CMS"
'    End If
' '   txtTotalEarned
'
'End Sub
Private Sub cmdInsertTotalEarned_Click()
    If Form_frmClientLedger.Closed = True Then
        Me.Total_Earned_Fee = Me.txtTotalEarned
    Else
        MsgBox "File Must be Closed to Insert Total Earned Fee.", , "TB CMS"
    End If
End Sub