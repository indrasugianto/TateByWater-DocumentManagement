' Component: Form_frmSourceAnalytics
' Type: document
' Lines: 103
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

Private Sub cmbClearFilter_Click()
    Call FilterClear
End Sub

Private Sub chkNonZero_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbCodeval_AfterUpdate()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbReferral_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdFilter_Click()
    Call FilterMe
End Sub

Private Sub cmdHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub

Private Sub txtIndRef_AfterUpdate()
    Call FilterMe
End Sub

Private Sub txtMatter_AfterUpdate()
    Call FilterMe
End Sub


Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbCodeVal) Then strSQL = strSQL & " AND CodeVal = '" & Me.cmbCodeVal & "'"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbReferral) Then strSQL = strSQL & " AND Referral = '" & Me.cmbReferral & "'"
    If Not IsNull(Me.txtMatter) Then strSQL = strSQL & " AND Matter_type like '*" & Me.txtMatter & "*'"
    If Not IsNull(Me.txtIndRef) Then strSQL = strSQL & " AND Matter_type like '*" & Me.txtIndRef & "*'"
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND ClsDate >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND ClsDate <= #" & Me.txtTo & "#"
    If Not IsNull(Me.chkNonZero) Then
        If chkNonZero Then
            strSQL = strSQL & " AND Expr2 <>0 "
        Else
            strSQL = strSQL & " AND Expr2 = 0 "
        End If
    End If
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub
Sub FilterClear()
   
    Me.cmbClients = Null
    Me.cmbCodeVal = Null
    Me.cmbOrigAtty = Null
    Me.txtMatter = Null
    Me.cmbReferral = Null
    Me.txtIndRef = Null
    Me.txtFrom = Null
    Me.txtTo = Null
    Me.chkNonZero = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub