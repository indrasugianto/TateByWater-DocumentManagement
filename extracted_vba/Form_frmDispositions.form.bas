' Component: Form_frmDispositions
' Type: document
' Lines: 143
' ============================================================

Option Compare Database

Private Sub CaseNo_Click()
    On Error GoTo ErrHandler_CaseNo_Click
    IsDisableEvents = True
    If CurrentProject.AllForms("frmClientLedger").IsLoaded Then
        DoCmd.Close acForm, "frmClientLedger", acSaveNo
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.CaseID, , , Me.CaseID
        Forms("frmClientLedger").SetFocus
    Else
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.CaseID, , , Me.CaseID
        Forms("frmClientLedger").SetFocus
    End If
ErrHandler_CaseNo_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub chkNGDis_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkNp_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkPISet_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkTrial_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbCodeval_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbCourt_AfterUpdate()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
   
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.cmbCodeVal) Then strSQL = strSQL & " AND CodeVal = '" & Me.cmbCodeVal & "'"
    If Not IsNull(Me.cmbCourt) Then strSQL = strSQL & " AND Court = '" & Me.cmbCourt & "'"
    If Not IsNull(Me.txtMatter) Then strSQL = strSQL & " AND Matter_Type like '*" & Me.txtMatter & "*'"
    If Not IsNull(Me.chkTrial) Then strSQL = strSQL & " AND Trial = " & IIf(Me.chkTrial, 1, 0)
    If Not IsNull(Me.chkNGDis) Then strSQL = strSQL & " AND [Not Guilty Dismissed] = " & IIf(Me.chkNGDis, 1, 0)
    If Not IsNull(Me.chkNp) Then strSQL = strSQL & " AND [Entire np] = " & IIf(Me.chkNp, 1, 0)
    If Not IsNull(Me.txtDispo) Then strSQL = strSQL & " AND Disposition like '*" & Me.txtDispo & "*'"
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND Dispo_Date >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND Dispo_Date <= #" & Me.txtTo & "#"
    If Not IsNull(Me.chkPISet) Then
        If chkPISet Then
            strSQL = strSQL & " AND [PI Settlement Amount] >0 "
        Else
            strSQL = strSQL & " AND [PI Settlement Amount] =0 "
        End If
    End If

    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
    
End Sub
Sub FilterClear()
  
Me.cmbClients = Null
Me.cmbAssoc = Null
Me.cmbOrigAtty = Null
Me.cmbCodeVal = Null
Me.cmbCourt = Null
Me.txtMatter = Null
Me.chkTrial = Null
Me.chkNGDis = Null
Me.chkNp = Null
Me.txtDispo = Null
Me.txtFrom = Null
Me.txtTo = Null
Me.chkPISet = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdFilter_Click()
    Call FilterMe
End Sub

Private Sub cmdHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub

'Sub FilterClear()
'
'Me.cmbClients = Null
'Me.cmbAssoc = Null
'Me.CmbOrigAtty = Null
'Me.chkCostHold = Null
'Me.chkAR = Null
'Me.chkUncashed = Null
'Me.chkUnclearedDep = Null
'Me.chkEarned = Null
'Me.chkCostReimb = Null
'
'
'    Me.filter = ""
'    Me.FilterOn = False
'End Sub

Private Sub txtDispo_AfterUpdate()
    If Not IsNull(txtDispo) Then Call FilterMe
End Sub

Private Sub txtMatter_AfterUpdate()
    If Not IsNull(txtMatter) Then Call FilterMe
End Sub