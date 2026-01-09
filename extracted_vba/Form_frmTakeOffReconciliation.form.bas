' Component: Form_frmTakeOffReconciliation
' Type: document
' Lines: 398
' ============================================================

Option Compare Database

'Private Sub cmdInsertIntoTA_Click()
'    If Me.txtEarned > 0 Then
'        If IsNull(txtTakeOffDate) Then
'            MsgBox "Please input the date"
'            Exit Sub
'        End If
'        Dim strSQL As String
'        strSQL = "Insert into [Trust Account] (CaseID, Credit, TDate, TMatter)"
'        strSQL = strSQL & " values (" & Me.CaseID & ", " & Me.txtEarned & ", #" & Me.txtTakeOffDate & "#, 'Earned Fee TO')"
'        Debug.Print strSQL
'        CurrentDb.Execute strSQL
'    End If
'
'    If Me.txtCostReimb > 0 Then
'        If IsNull(txtTakeOffDate) Then
'            MsgBox "Please input the date"
'            Exit Sub
'        End If
'        Dim strSQL3 As String
'        strSQL3 = "Insert into [Trust Account] (CaseID, Credit, TDate, TMatter)"
'        strSQL3 = strSQL3 & " values (" & Me.CaseID & ", " & Me.txtCostReimb & ", #" & Me.txtTakeOffDate & "#, 'Cost Reimb TO')"
'        Debug.Print strSQL3
'        CurrentDb.Execute strSQL3
'    End If
'
'    'Take off table:
'    Dim strSQL2 As String
'    strSQL2 = "Insert into [tblTakeOff] "
'    strSQL2 = strSQL2 & "        (CaseID, TakeOffDate, EarlyEarned, TOEarned, CostReimb, CBHRev, MKRev, MTRev, CBHCom, MTCom, KBCom, MKCom, EarlyEarnedTr, TOEarnedTr, CostReimbTr)"
'    strSQL2 = strSQL2 & " values (" & Me.CaseID & ", #" & Me.txtTakeOffDate & "#, " & Me.txtEarlyEarned & ", " & Me.txtEarned & ", " & _
'                        Me.txtCostReimb & ", " & Me.txtCBHRev & ", " & Me.txtMKRev & ", " & Me.txtMTRev & ", " & Me.txtCBHCom & ", " & _
'                        Me.txtMTCom & ", " & Me.txtKDBCom & ", " & Me.txtMKCom & ", " & Me.ChkEarlyEarnedTR & ", " & Me.chkTOEarnedTr & ", " & Me.chkCostReimbTr & ")"

Private Sub cmdAttyReport_Click()
    DoCmd.Close acReport, "Client_Trust_Accounts_for_PreTake_Off", acSaveYes
    DoCmd.OpenReport "Client_Trust_Accounts_for_PreTake_Off", acViewPreview, "", "Orig_Atty='" & Me.cmbAttyPTO & "'"
    DoCmd.Close acForm, "Client_Trust_Accounts_for_PreTake_Off"
End Sub

'    Debug.Print strSQL2
'    CurrentDb.Execute strSQL2
'
'    'clear fields:
'    Me.txtEarlyEarned = Null
'    Me.txtEarned = Null
'    Me.txtCostReimb = Null
'    Me.txtCBHRev = Null
'    Me.txtMKRev = Null
'    Me.txtCBHCom = Null
'    Me.txtMTCom = Null
'    Me.txtKDBCom = Null
'    Me.txtMKCom = Null
'    Me.txtMTRev = Null
'    Me.ChkEarlyEarnedTR = Null
'    Me.chkTOEarnedTr = Null
'    Me.chkCostReimbTr = Null
'
'    MsgBox "Done"
'
'End Sub
Private Sub Form_Load()
    
    'Me.txtJMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=11946")
    'Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
    Me.txtSMTrust = DLookup("AvailBalance", "qryTakeOff", "CaseID=12664")
    Me.txtRLFDeeds = DLookup("AvailBalance", "qryTakeOff", "CaseID=23437")
    Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Me.txtSMTrust
    'Me.txtSumOfBalances = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances") - Me.txtJMTrust - Me.txtSMTrust
    Me.txtSumOfUncashed = DLookup("SumOfCredit", "qryReconciliation_sumOfCredit")  'Not cashed! please check qyery !
    Me.txtSumOfUnclearDeposits = DLookup("SumOfDebit", "qryReconciliation_sumOfUnclearedDeposits")
    Me.txt_WF_CHK = Me.txtSumOfBalances + Nz(Me.txtSumOfUncashed, 0) - Nz(Me.txtSumOfUnclearDeposits, 0) 'expwftrust+uncash-uncleardep
    Me.txtTotalMinusInputBalance = Me.txt_WF_CHK - Me.txt_WF_balance_trust_amount
    'txtJMTrust = DLookup("SumOfBalance", "qryReconciliation_sumOfBalances")
    'txtActualJMTrust 'input field
    Me.txtWFTrustAndJMSMTrust = Me.txtSumOfBalances + Me.txtSMTrust
    'Me.txtWFTrustAndJMSMTrust = Me.txtSumOfBalances + Me.txtJMTrust + Me.txtSMTrust
    
    Me.cmdInsertData.Enabled = True
'    If fncReconciliationExists Then
'        cmdInsertData.Enabled = False
'    Else
'        cmdInsertData.Enabled = True
'    End If
End Sub

Sub checkReconciledCondition()
'    If chkReconciled.value = -1 And fncReconciliationExists = False Then
'        Me.cmdInsertData.Enabled = True
'    Else
'        Me.cmdInsertData.Enabled = False
'    End If
End Sub

Private Sub txtTKButton_Click()
    
    Dim strSQLField As String
    Dim strSQLFinal As String
    If Nz(Me.SumOfTotal, 0) = 0 Then
        Exit Sub
    End If
    
    maxOrderNr = DMax("OrderNr", "[Matter and AR]", "CaseID=" & Me.CaseID)
    
    If Me.AvailBalance >= Me.SumOfTotal Then
        strSQLField = "StatementLessTrust"
    ElseIf Me.AvailBalance < Me.SumOfTotal Then
        strSQLField = "InvoiceExceedsTrust"
        
        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL
        
    ElseIf Me.AvailBalance <= 0 Then
        strSQLField = "InvoiceTotal"
        
        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL
        
    End If
    
    strSQLFinal = "Update [TB Time Keeping] set " & _
                    "TrustatClose = " & Me.AvailBalance & ", " & _
                    strSQLField & "=-1, " & _
                    "[Bill Closed]=-1, " & _
                    "[Bill Closed Date] = #" & Format(Date, "yyyy-MM-dd") & "#," & _
                    "TKLocked=-1 " & _
                    " where Bill_ID=" & Me.Bill_ID
                    
    CurrentDb.Execute strSQLFinal
    
    Call Form_frmMatter.reorderByDateMatter(Me.CaseID)
    'Dim strBookmark As String
    'strBookmark = Me.Bookmark
    
    Form_frmMatter.Requery
    Me.Requery
    
    'Me.Bookmark = strBookmark
    
    'Fields to Update: Bill Closed, Bill Closed Date, TKLocked, Trustat Close, InvoiceTotal OR InvoiceExceedsTrust OR StatementLessTrust
    'From            : -1         , Date ()         , -1      , AvailBalance,  -1
    
    'strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID

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

Private Sub chkAdvFee_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkAR_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkCostHold_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkNonZero_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkOpenTK_AfterUpdate()
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

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub chkReconciled_Click()
    Call checkReconciledCondition
End Sub

Private Sub cmdRequery_Click()
    Form_frmTakeOffReconciliation.Requery
End Sub

'Private Sub txt_total_minus_inputbalance_AfterUpdate()
'    Call checkReconciledCondition
'End Sub

Private Sub txt_WF_balance_trust_amount_AfterUpdate()
    Me.txtTotalMinusInputBalance = Me.txt_WF_CHK - Me.txt_WF_balance_trust_amount
    Call checkReconciledCondition
End Sub

Private Sub cmdInsertData_Click()

    'simple date check:
'    ExistingCaseID = DLookup("CaseID", "qryTakeOff", "TakeOffDate = #" & Date & "#")
'    If ExistingCaseID > 0 Then
'        MsgBox "Reconciliation for this date exists in the database."
'        Exit Sub
'    End If
    
    'month check:
    'strFilter = "MonthOnly = " & Month(Date) & " and YearOnly = " & Year(Date)
    'ExistingCaseID = DLookup("TakeOffMonthID", "qry_takeOff_year_month", strFilter)
    'If ExistingCaseID > 0 Then
    
'    If fncReconciliationExists Then
'
'        MsgBox "Reconciliation for this month exists in the database.", vbExclamation, "TB CMS"
'        Exit Sub
'    Else
        
        Dim strSQL As String
        
        strSQL = "Insert into [tblTakeOffMonth] "
        strSQL = strSQL & "(TakeOffDate, [WF Balance], LastWF, SumUncashed, SumUncleared, WFActual, CombinedTrust, SomBalance, SomActual,"
        strSQL = strSQL & "RLFDeedFees, ReconcileValue, AccReconciled, WFplusuncashed)"
        strSQL = strSQL & " values (#" & Date & "#, " & Nz(txtSumOfBalances, 0) & ", '" & Nz(Me.txtLastWF, "") & "', " & Nz(txtSumOfUncashed, 0) & ", "
        strSQL = strSQL & Nz(txtSumOfUnclearDeposits, 0) & ", " & Nz(txt_WF_balance_trust_amount, 0) & ", " & Nz(txtWFTrustAndJMSMTrust, 0) & ", "
        strSQL = strSQL & Nz(Me.txtSMTrust, 0) & ", " & Nz(Me.txtActualSMTrust, 0) & ", " & Nz(Me.txtRLFDeeds, 0) & "," & Nz(Me.txtTotalMinusInputBalance, 0) & ", "
        strSQL = strSQL & Nz(chkReconciled.value, 0) & ", " & Nz(Me.txt_WF_CHK, 0) & ")"
        
    
        'Insert into "Take Off Month" table:
        'strSQL = "Insert into [tblTakeOffMonth] "
        'strSQL = strSQL & "        (TakeOffDate, [WF Balance], LastWF, SumUncashed, SumUncleared, WFActual, CombinedTrust, " & _
                                   "DaleBalance, DaleActual, SomBalance, SomActual, ReconcileValue, AccReconciled, WFplusuncashed)"
        'strSQL = strSQL & " values (#" & Date & "#, " & Nz(txtSumOfBalances, 0) & ", '" & Nz(Me.txtLastWF) & "', " & Nz(txtSumOfUncashed, 0) & ", " & _
                                    Nz(txtSumOfUnclearDeposits, 0) & ", " & Nz(txt_WF_balance_trust_amount, 0) & ", " & Nz(txtWFTrustAndJMSMTrust, 0) & ", " & _
                                    Me.txtJMTrust & ", " & Me.txtActualJMTrust & ", " & Me.txtSMTrust & ", " & Me.txtActualSMTrust & "," & Me.txtTotalMinusInputBalance & "," & _
                                    chkReconciled.value & ", " & Nz(Me.txt_WF_CHK, 0) & ")"
                          
        Debug.Print strSQL
        CurrentDb.Execute strSQL
        
        MaxID_tblTakeOffMonth = DMax("TakeOffMonthID", "tblTakeOffMonth")
        
        Dim rs As Recordset
        'Set rs = CurrentDb.OpenRecordset("select * from qryTakeOff where AvailBalance > 0 ") 'take only non-empty records!   'Me.RecordsetClone
        Set rs = Me.RecordsetClone 'CurrentDb.OpenRecordset("select * from qryTakeOff")
        Do Until rs.EOF
'
'            'Take off table:
            strSQL = "Insert into [tblTakeOff] "
            strSQL = strSQL & "        (CaseID, TakeOffMonthID, AvailBalance, TotalUnCashedChks, TotalUnclearedDeps, AdvCostBal, AdvFeeBal, TotalAdvancedAR, TotalHourlyOuts, OpenTK)"
            strSQL = strSQL & " values (" & rs("CaseID") & ", " & MaxID_tblTakeOffMonth & ", " & Nz(rs("AvailBalance"), 0) & ", " & Nz(rs("SumOfUncashedChecks"), 0) & ", " & _
                                        Nz(rs("SumOfUnclearedDeposits"), 0) & ", " & Nz(rs("AdvCostBalance"), 0) & ", " & Nz(rs("AdvLegalBalance"), 0) & ", " & Nz(rs("CostHoldBalance"), 0) & ", " & Nz(rs("SumOfTotal"), 0) & ", '" & Nz(rs("IANumber"), "") & "')"
'                                       Nz(rs("SumOfUnclearedDeposits"), 0) & ", " & Nz(txtAdvancedCostBalance, 0) & ", " & Nz(txtAdvancedFeesBalance, 0) & ", " & Nz(txtCostResBalance, 0) & ", " & Nz(rs("SumOfTotal"), 0) & ", '" & Nz(rs("IANumber"), "") & "')"
'Need to include IAnumber into OpenTK
'                                        'if we run into a trouble inserting the records, please check the field TotalHourlyOuts and the text box: Nz(rs("SumOfTotal")
                                         
                                                    
            'controls:
            'txt_total_minus_inputbalance
            'txtJMTrust
            'txtActualJMTrust

            Debug.Print strSQL
            CurrentDb.Execute strSQL
            
            rs.MoveNext
        Loop
        
        
        
        MsgBox "Done", , "TB CMS"
    'End If

    'Me.cmdInsertData.Enabled = False
    
End Sub

Function fncReconciliationExists() As Boolean
    strFilter = "MonthOnly = " & Month(Date) & " and YearOnly = " & Year(Date)
    ExistingID = DLookup("TakeOffMonthID", "qry_takeOff_year_month", strFilter)

    If ExistingID > 0 Then
        fncReconciliationExists = True
    Else
        fncReconciliationExists = False
    End If
End Function

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
'    If Not IsNull(Me.chkAR) Then
'        If chkAR Then
'            strSQL = strSQL & " AND txtAdvancedCostBalance >0 "
'        Else
'            strSQL = strSQL & " AND txtAdvancedCostBalance =0 "
'        End If
'    End If
'    If Not IsNull(Me.chkAdvFee) Then
'        If chkAR Then
'            strSQL = strSQL & " AND txtAdvancedFeesBalance >0 "
'        Else
'            strSQL = strSQL & " AND txtAdvancedFeesBalance =0 "
'        End If
'    End If
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
      If Not IsNull(Me.chkOpenTK) Then
        If chkOpenTK Then
            strSQL = strSQL & " AND IANumber like '*'"
        Else
            strSQL = strSQL & " AND IANumber not like '*'"
        End If
    End If
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()
   
    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.chkUnclearedDep = Null
    Me.chkOpenTK = Null
'    Me.chkAR = Null
    Me.ChkUncashed = Null
    Me.chkCostHold = Null
'    Me.chkAdvFee = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

