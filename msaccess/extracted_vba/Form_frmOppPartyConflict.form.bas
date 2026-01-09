' Component: Form_frmOppPartyConflict
' Type: document
' Lines: 68
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

Private Sub txtOPartyDOB_AfterUpdate()
    If Not IsNull(txtOPartyDOB) Then Call FilterMe
End Sub

Private Sub txtOPartyFirst_AfterUpdate()
    If Not IsNull(txtOPartyFirst) Then Call FilterMe
End Sub



Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.TxtOPartyLast) Then strSQL = strSQL & " AND OPartyLast like '*" & Me.TxtOPartyLast & "*'"
    If Not IsNull(Me.txtOPartyFirst) Then strSQL = strSQL & " AND OPartyFirst like '*" & Me.txtOPartyFirst & "*'"
    If Not IsNull(Me.txtOPartyDOB) Then strSQL = strSQL & " AND OPartyDOB like '" & Me.txtOPartyDOB & "*'"
    If Not IsNull(Me.txtMatterClient) Then strSQL = strSQL & " AND Matter_type like '*" & Me.txtMatterClient & "*'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()
    'clear controls:
    Me.TxtOPartyLast = Null
    Me.txtOPartyFirst = Null
    Me.txtOPartyDOB = Null
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

Private Sub TxtOPartyLast_AfterUpdate()
    If Not IsNull(TxtOPartyLast) Then Call FilterMe
End Sub