' Component: Form_frmClientsConflict
' Type: document
' Lines: 66
' ============================================================

Option Compare Database

Private Sub CaseNum_Click()
On Error GoTo ErrHandler_CaseNum_Click
    'IsDisableEvents = True
    If CurrentProject.AllForms("frmClientLedger").IsLoaded Then
        DoCmd.Close acForm, "frmClientLedger", acSaveNo
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.CaseID, , , Me.CaseID
        Forms("frmClientLedger").SetFocus
    Else
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.CaseID, , , Me.CaseID
        Forms("frmClientLedger").SetFocus
    End If
ErrHandler_CaseNum_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description

End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub txtDOBClient_AfterUpdate()
    If Not IsNull(txtDOBClient) Then Call FilterMe
End Sub

Private Sub txtFirstClient_AfterUpdate()
    If Not IsNull(txtFirstClient) Then Call FilterMe
End Sub

Private Sub txtLastClient_AfterUpdate()
    If Not IsNull(txtLastClient) Then Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.txtLastClient) Then strSQL = strSQL & " AND Last_Name like '*" & Me.txtLastClient & "*'"
    If Not IsNull(Me.txtFirstClient) Then strSQL = strSQL & " AND First_Name like '*" & Me.txtFirstClient & "*'"
    If Not IsNull(Me.txtDOBClient) Then strSQL = strSQL & " AND DOB like '" & Me.txtDOBClient & "*'"
    If Not IsNull(Me.txtMatterClient) Then strSQL = strSQL & " AND Matter_type like '*" & Me.txtMatterClient & "*'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()
    'clear controls:
    Me.txtLastClient = Null
    Me.txtFirstClient = Null
    Me.txtDOBClient = Null
    Me.txtMatterClient = Null
'    Me.chkPastDue = Null
'    Me.chkNoBalance = Null
'    Me.chkNonZero = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub txtMatterClient_AfterUpdate()
    If Not IsNull(txtMatterClient) Then Call FilterMe
End Sub