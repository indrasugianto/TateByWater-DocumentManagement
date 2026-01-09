' Component: Form_frm_advanced_payments
' Type: document
' Lines: 113
' ============================================================

Option Compare Database

Private Sub chkNonZero_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClearFilter_Click()
    Call FilterClear
End Sub

'Private Sub Case_Letter_AfterUpdate()
'    Call FilterMe
'End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND [Matter and AR].Date2 >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND [Matter and AR].Date2 <= #" & Me.txtTo & "#"
    
    If Not IsNull(Me.cmbCodeVal) Then strSQL = strSQL & " AND CodeVal = '" & Me.cmbCodeVal & "'"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '" & Me.txtClient & "*'"
    
    'If Not IsNull(Me.chkPastDue) Then strSQL = strSQL & " AND [Past Due] = " & Me.chkPastDue
    'If Not IsNull(Me.chkNoBalance) Then strSQL = strSQL & " AND chkBalanceDue = " & Me.chkNoBalance
    
'    If Not IsNull(Me.chkNonZero) Then
'       If chkNonZero Then
'            strSQL = strSQL & " AND TrustAcctBalance <>0 "
'        Else
'            strSQL = strSQL & " AND TrustAcctBalance =0 "
'        End If
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub cmbclose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmbCodeVal_Click()
    Call FilterMe
End Sub

Sub FilterClear()
    'clear controls:
    Me.cmbClients = Null
'    Me.chkPastDue = Null
'    Me.chkNoBalance = Null
'    Me.chkNonZero = Null
    
    Me.txtFrom = Null
    Me.txtTo = Null
    Me.cmbCodeVal = Null
    Me.cmbOrigAtty = Null
    Me.txtClient = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmbFilter_Click()
    Call FilterMe
End Sub

Private Sub cmbHome_Click()
DoCmd.openform "frmhome", acNormal
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub FileNumber_Click()
On Error GoTo ErrHandler_CaseNum_Click
    IsDisableEvents = True
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

Private Sub cmdFilter_Click()
    Call FilterMe
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub