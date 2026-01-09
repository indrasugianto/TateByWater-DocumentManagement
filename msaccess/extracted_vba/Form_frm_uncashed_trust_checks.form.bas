' Component: Form_frm_uncashed_trust_checks
' Type: document
' Lines: 57
' ============================================================

Option Compare Database

Private Sub CaseNum_Click()
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

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.txtCheck) Then strSQL = strSQL & " AND CheckNumber like '" & Me.txtCheck & "*'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Sub FilterClear()
    'clear controls:
    Me.cmbClients = Null
    Me.txtCheck = Null
'    Me.chkPastDue = Null
'    Me.chkNoBalance = Null
'    Me.chkNonZero = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub txtCheck_AfterUpdate()
    If Not IsNull(txtCheck) Then Call FilterMe
End Sub