' Component: Form_frmTakeOffSubForm3
' Type: document
' Lines: 261
' ============================================================

Option Compare Database

Private Sub chkAR_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkCostHold_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkCostReimb_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkEarned_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkUncashed_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkUnclearedDep_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
   
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.chkAR) Then
        If chkAR Then
            strSQL = strSQL & " AND SumOfAdvancedAR >0 "
        Else
            strSQL = strSQL & " AND SumOfAdvancedAR =0 "
        End If
    End If
    If Not IsNull(Me.chkCostHold) Then
        If chkCostHold Then
            strSQL = strSQL & " AND CostHold >0 "
        Else
            strSQL = strSQL & " AND CostHold =0 "
        End If
    End If
    If Not IsNull(Me.ChkUncashed) Then
        If ChkUncashed Then
            strSQL = strSQL & " AND SumOfUncashedChecks >0 "
        Else
            strSQL = strSQL & " AND SumOfUncashedChecks =0 "
        End If
    End If
    If Not IsNull(Me.chkUnclearedDep) Then
        If chkUnclearedDep Then
            strSQL = strSQL & " AND SumOfUnclearedDeposits >0 "
        Else
            strSQL = strSQL & " AND SumOfUnclearedDeposits =0 "
        End If
    End If
    If Not IsNull(Me.chkEarned) Then
        If chkEarned Then
            strSQL = strSQL & " AND TOEarned >0 "
        Else
            strSQL = strSQL & " AND TOEarned =0 "
        End If
    End If
    If Not IsNull(Me.chkCostReimb) Then
        If chkCostReimb Then
            strSQL = strSQL & " AND CostReimb >0 "
        Else
            strSQL = strSQL & " AND CostReimb =0 "
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
Me.chkCostHold = Null
Me.chkAR = Null
Me.ChkUncashed = Null
Me.chkUnclearedDep = Null
Me.chkEarned = Null
Me.chkCostReimb = Null

    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub CaseNum_Click()
On Error GoTo ErrHandler_CaseNum_Click
Dim intCaseID As Integer

    'intCaseID = Split(Me.FileNumber, "-")(1)
    intCaseID = Me.txtCaseID
    

    IsDisableEvents = True
    If CurrentProject.AllForms("frmClientLedger").IsLoaded Then
        DoCmd.Close acForm, "frmClientLedger", acSaveNo
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & intCaseID, , , intCaseID
        Forms("frmClientLedger").SetFocus
    Else
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & intCaseID, , , intCaseID
        Forms("frmClientLedger").SetFocus
    End If
ErrHandler_CaseNum_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

'Private Sub cmdInsertIntoTA_Click()
''    If Me.txtEarned > 0 Then
'''        If IsNull(txtTakeOffDate) Then
'''            MsgBox "Please input the date"
'''            Exit Sub
'''        End If
''        Dim strSQL As String
''        strSQL = "Insert into [Trust Account] (CaseID, Credit, TDate, TMatter)"
''        strSQL = strSQL & " values (" & _
''                          Me.txtCaseID & ", " & _
''                          Nz(Me.txtEarned, 0) & ", #" & _
''                          Form_frmTakeOff.txtTakeOffDate & "#, 'Earned Fee')"
''        Debug.Print strSQL
''        CurrentDb.Execute strSQL
''        i = 1
''    End If
'
'    If Me.txtAdvEarned > 0 Then
'        Dim strSQL As String
'        strSQL = "Insert into [Trust Account] (CaseID, Credit, TDate, TMatter, AdvFee)"
'        strSQL = strSQL & " values (" & _
'                          Me.txtCaseID & ", " & _
'                          Nz(Me.txtAdvEarned, 0) & ", #" & _
'                          Form_frmTakeOff.txtTakeOffDate & "#, 'Earned Fee (ADV)', -1)"
'        Debug.Print strSQL
'        CurrentDb.Execute strSQL
'        i = 1
'    End If
'
'    If Me.txtRemEarned > 0 Then
'        Dim strSQL4 As String
'        strSQL4 = "Insert into [Trust Account] (CaseID, Credit, TDate, TMatter)"
'        strSQL4 = strSQL4 & " values (" & _
'                          Me.txtCaseID & ", " & _
'                          Nz(Me.txtRemEarned, 0) & ", #" & _
'                          Form_frmTakeOff.txtTakeOffDate & "#, 'Earned Fee')"
'        Debug.Print strSQL4
'        CurrentDb.Execute strSQL4
'        i = 1
'    End If
'
'    If Me.txtCostReimb > 0 Then
''        If IsNull(txtTakeOffDate) Then
''            MsgBox "Please input the date"
''            Exit Sub
''        End If
'        Dim strSQL3 As String
'        strSQL3 = "Insert into [Trust Account] (CaseID, Credit, TDate, TMatter, CostReimb)"
'        strSQL3 = strSQL3 & " values (" & Me.txtCaseID & ", " & Me.txtCostReimb & ", #" & Form_frmTakeOff.txtTakeOffDate & "#, 'Cost Reimb', -1)"
'        Debug.Print strSQL3
'        CurrentDb.Execute strSQL3
'        i = 1
'        answer = MsgBox("Do you want to insert Cost Reimbursement Credit in AR?", vbYesNo, "TB CMS")
'        If answer = vbYes Then
'        Dim strSQL5 As String
'        strSQL5 = "Insert into [Matter and AR] (CaseID, Payment, Date2, Pay_Outlay)"
'        strSQL5 = strSQL5 & " values (" & Me.txtCaseID & ", " & Me.txtCostReimb & ", #" & Form_frmTakeOff.txtTakeOffDate & "#, 'Cost Reimb from Trust')"
'        Debug.Print strSQL5
'        CurrentDb.Execute strSQL5
'        End If
'
'    End If
'
'    If i >= 1 Then
'        Me.Form.Recordset.Edit
'        Me.Form.Recordset("InsertedTrust") = -1
'        Me.Form.Recordset.Update
'    End If
'
''    Select Case i
''        Case 1
''            MsgBox "Added an Earned Fee"
''        Case 2
''            MsgBox "Added a Cost Reimbursement"
''        Case 3
''            MsgBox "Added Earned Fee and Cost Reimbursement"
''        Case Else
''            MsgBox "Please input an Earned Fee or Cost Reimbursement"
''    End Select
'    'Me.Requery
'End Sub

'Private Sub Form_Current()
'    Me.cmdInsertIntoTA.Enabled = Not Me.Form.Recordset("InsertedTrust")
'End Sub

'Private Sub Detail_Paint()
'    Me.cmdInsertIntoTA.Enabled = Not Nz(Me.InsertedTrust, -1) 'Not Me.Form.Recordset("InsertedTrust")
'End Sub

'Private Sub txtTrustButton_Click()
'    If Me.InsertedTrust = 0 Then
'        Call cmdInsertIntoTA_Click
'    End If
'End Sub

'Private Sub Form_Load() NOT WORKING!
'    txtTrustButton.MousePointer = 99
'    'txtTrustButton.MouseIcon = path("d:\! Jobs\Paul Mickelsen - Access Lawyer Application\Link.cur")
'    txtTrustButton.MouseIcon = "d:\! Jobs\Paul Mickelsen - Access Lawyer Application\Link.cur"
'End Sub

Private Sub txtEarned_AfterUpdate()
If TOEarned > AvailBalance Then
         MsgBox "Earned Cannot be Greater than Available Trust", vbCritical, "TB CMS"
        Me.Undo
Else
Dim sFileNumber As String

        sFileNumber = Me.FileNumber
        
        bk = Me.CurrentRecord
        DoCmd.RunCommand acCmdSaveRecord
        Me.Requery
        Me.Recordset.FindFirst "FileNumber = " & pcaAddQuotes(sFileNumber)
        Me.Refresh
End If
End Sub

Private Sub txtCostReimb_AfterUpdate()
    'Form_frmTakeOff.Requery
End Sub
