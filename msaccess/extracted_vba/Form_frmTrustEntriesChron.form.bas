' Component: Form_frmTrustEntriesChron
' Type: document
' Lines: 220
' ============================================================

Option Compare Database

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    'If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    'If Not IsNull(Me.chkPastDue) Then strSQL = strSQL & " AND [Past Due] = " & Me.chkPastDue
    'If Not IsNull(Me.chkNoBalance) Then strSQL = strSQL & " AND chkBalanceDue = " & Me.chkNoBalance
    
'    If Not IsNull(Me.chkNonZero) Then
'        If chkNonZero Then
'            strSQL = strSQL & " AND BalRet >0 "
'        Else
'            strSQL = strSQL & " AND BalRet =0 "
'        End If
'    End If
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND TDate >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND TDate <= #" & Me.txtTo & "#"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmdWFReconcile) Then strSQL = strSQL & " AND Reconciled = " & IIf(Me.cmdWFReconcile, 1, 0)
    If Not IsNull(Me.ChkUncashed) Then strSQL = strSQL & " AND CheckCashed = " & IIf(Me.ChkUncashed, 1, 0)
    If Not IsNull(Me.chkUnclearedDep) Then strSQL = strSQL & " AND DepCleared = " & IIf(Me.chkUnclearedDep, 1, 0)
    If Not IsNull(Me.txtCheck) Then strSQL = strSQL & " AND CheckNumber like '*" & Me.txtCheck & "*'"
    If Not IsNull(Me.txtDescription) Then strSQL = strSQL & " AND TMatter like '*" & Me.txtDescription & "*'"
    If Not IsNull(Me.txtWith) Then strSQL = strSQL & " AND Credit like '" & Me.txtWith & "*'"
    If Not IsNull(Me.txtDep) Then strSQL = strSQL & " AND Debit like '" & Me.txtDep & "*'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '*" & Me.txtClient & "*'"
'    If Not IsNull(Me.chkChks) Then strSQL = strSQL & " AND CheckNumber like '*' & Me.chkChks & '*'"
    If Not IsNull(Me.chkNonZeroD) Then
        If chkNonZeroD Then
            strSQL = strSQL & " AND Debit <>0 "
        Else
            strSQL = strSQL & " AND Debit = 0 "
        End If
    End If
    If Not IsNull(Me.chkNonZeroW) Then
        If chkNonZeroW Then
            strSQL = strSQL & " AND Credit <>0 "
        Else
            strSQL = strSQL & " AND Credit = 0 "
        End If
    End If
    If Not IsNull(Me.chkChks) Then
        If chkChks Then
            strSQL = strSQL & " AND Checknumber like '*'"
        Else
            strSQL = strSQL & " AND Checknumber not like '*'"
        End If
    End If
    
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

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

Private Sub chkChks_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkNonZeroD_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkNonZeroW_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkUncashed_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkUnclearedDep_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmd35D_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_35D", acViewPreview
End Sub
Private Sub cmd35w_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_35W", acViewPreview
End Sub
Private Sub cmd35_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_35", acViewPreview
End Sub
Private Sub cmd65_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_65", acViewPreview
End Sub
Private Sub cmd65D_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_65D", acViewPreview
End Sub
Private Sub cmd65W_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_65W", acViewPreview
End Sub
Private Sub cmd95_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_95", acViewPreview
End Sub
Private Sub cmd95D_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_95D", acViewPreview
End Sub
Private Sub cmd95W_Click()
    DoCmd.OpenReport "rpt_Trust_Chron_95W", acViewPreview
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

'Sub FilterClear()
'    'clear controls:
'    'Me.cmbClients = Null
''    Me.chkPastDue = Null
''    Me.chkNoBalance = Null
''    Me.chkNonZero = Null
'
'    Me.filter = ""
'    Me.FilterOn = False
'End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdFilter_Click()
    FilterMe
End Sub

Sub FilterClear()
    'clear controls:
    Me.cmbClients = Null
'    Me.chkPastDue = Null
'    Me.chkNoBalance = Null
'    Me.chkNonZero = Null
    Me.txtFrom = Null
    Me.txtTo = Null
    Me.cmbOrigAtty = Null
    Me.cmdWFReconcile = Null
    Me.ChkUncashed = Null
    Me.chkUnclearedDep = Null
    Me.txtCheck = Null
    Me.txtDescription = Null
    Me.txtWith = Null
    Me.txtDep = Null
    Me.chkNonZeroD = Null
    Me.chkNonZeroW = Null
    Me.txtClient = Null
    Me.chkChks = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdRequery_Click()
    If Me.Dirty = True Then Me.Dirty = False
    Me.Requery
    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Me.txtSMTrust
    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")
    Me.txt_WF_CHK = Me.txtSumOfBalances + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
End Sub

Private Sub cmdWFReconcile_AfterUpdate()
    Call FilterMe
End Sub

Private Sub Form_Load()
    'Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=11946")
    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Me.txtSMTrust
    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")
    Me.txt_WF_CHK = Me.txtSumOfBalances + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
End Sub

Private Sub txtCheck_AfterUpdate()
    If Not IsNull(txtCheck) Then Call FilterMe
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub

Private Sub txtDep_AfterUpdate()
    If Not IsNull(txtDep) Then Call FilterMe
End Sub

Private Sub txtDescription_AfterUpdate()
    If Not IsNull(txtDescription) Then Call FilterMe
End Sub

Private Sub txtWith_AfterUpdate()
    If Not IsNull(txtWith) Then Call FilterMe
End Sub