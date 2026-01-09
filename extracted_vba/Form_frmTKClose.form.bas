' Component: Form_frmTKClose
' Type: document
' Lines: 456
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

Private Sub chkAR_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkAROutstanding_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkCostHold_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkNonZero_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkHourlyOuts_AfterUpdate()
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
Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.chkCostHold) Then
        If chkCostHold Then
            strSQL = strSQL & " AND CostHold >0 "
        Else
            strSQL = strSQL & " AND CostHold =0 "
        End If
    End If
    If Not IsNull(Me.chkAROutstanding) Then
        If chkAROutstanding Then
            strSQL = strSQL & " AND txtAROutstanding >0 "
        Else
            strSQL = strSQL & " AND txtAROutstanding =0 "
        End If
    End If
'     If Not IsNull(Me.chkHourlyOuts) Then
'        If chkHourlyOuts Then
'            strSQL = strSQL & " AND SumofTotal >0 "
'        Else
'            strSQL = strSQL & " AND SumofTotal =0 "
'        End If
'    End If
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()
   
    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.chkAROutstanding = Null
'    Me.chkHourlyOuts = Null
    Me.chkCostHold = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub Command228_Click()
    Me.Requery
End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
 If Not Me.Dirty Then
   If (Count < 0) And (Me.CurrentRecord > 1) Then
     DoCmd.GoToRecord , , acPrevious
   ElseIf (Count > 0) And (Me.CurrentRecord <= Me.Recordset.RecordCount) Then
        On Error GoTo errFix
        DoCmd.GoToRecord , , acNext
   End If
 Else
   MsgBox "The record has changed. Save the current record before moving to another record.", , "TB CMS"
 End If
errFix:
    If Err.Description = "You can't go to the specified record." Then
    End If
End Sub

Private Sub txtTKButton_Click()
    
    Dim strSQLField As String
    Dim strSQLFinal As String
    'If Nz(Me.SumOfTotal, 0) = 0 Then
        'Exit Sub
    'End If
    
    maxOrderNr = DMax("OrderNr", "[Matter and AR]", "CaseID=" & Me.CaseID)
    
    'DEFAULT
    
    strSQLField = "InvoiceExceedsTrust"
    
    'STATEMENT (WORKS!)
    
    If Me.txtTrustafterTK > 0 And Me.txtAROutstanding <= 0 And Me.txtReplBalance <= 0 And Nz(Me.txtAdvInvoice, 0) <= 0 Then
        strSQLField = "StatementLessTrust" 'No advanced legal fee

    'TOTAL ADVANCE no REPLENISH (with no prior existing advanced fee balance)(WORKS!)

    ElseIf Me.txtAvailTrust <= 0 And Me.txtReplBalance <= 0 And Nz(Me.txtAdvancedCostBalance, 0) <= 0 And Nz(Me.txtAdvancedFeesBalance, 0) <= 0 And Nz(Me.txtCostResBalance, 0) <= 0 Then
        strSQLField = "InvoiceTotalAdvance" 'advanced legal fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
    
    'TOTAL ADVANCE with REPLENISH (no advanced costs or fees) (WORKS!)
        
     ElseIf Me.txtAvailTrust <= 0 And Me.txtReplBalance > 0 And Nz(Me.txtAdvancedCostBalance, 0) <= 0 And Nz(Me.txtAdvancedFeesBalance, 0) <= 0 And Nz(Me.txtCostResBalance, 0) <= 0 Then
        strSQLField = "InvoiceExceedsTrust" 'advanced legal fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges

    'TOTAL ADVANCE with REPLENISH (with advanced costs or fees)
    
    ElseIf Me.txtAvailTrust <= 0 And Me.txtReplBalance > 0 And Nz(Me.txtAdvancedCostBalance, 0) > 0 Or Nz(Me.txtAdvancedFeesBalance, 0) > 0 Or Nz(Me.txtCostResBalance, 0) > 0 Then
        strSQLField = "InvoiceAdvCostFee" 'advanced legal fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges


'    ElseIf Me.txtAvailTrust < Me.txtSumofTotal And Me.txtAvailTrust > 0 And Nz(Me.RetainerReimb, 0) = 0 Then
'        strSQLField = "InvoiceExceedsTrust" 'advanced legal fee
'
'        strSQL = "Insert into [Matter and AR] "
'            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
'            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", " & maxOrderNr + 1 & ")"
'        Debug.Print strSQL
'        CurrentDb.Execute strSQL

    'INVOICE (Advanced TK) but NO ADVANCED COSTS or HOLDS and NO REPLENISH (TKEx1)
    
    ElseIf Me.txtAvailTrust > 0 And Me.txtReplBalance <= 0 And Nz(Me.txtAdvancedCostBalance, 0) <= 0 And Nz(Me.txtAdvancedFeesBalance, 0) <= 0 And Nz(Me.txtCostResBalance, 0) <= 0 And Nz(Me.txtAdvInvoice, 0) > 0 Then
        strSQLField = "InvoiceExceedsTrust" 'advanced legal fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
     'INVOICE (No Advanced TK) but NO ADVANCED COSTS/Fees or HOLDS and YES REPLENISH (TKEx1)
     
     ElseIf Me.txtAvailTrust > 0 And Me.txtReplBalance > 0 And Nz(Me.txtAdvancedCostBalance, 0) <= 0 And Nz(Me.txtAdvancedFeesBalance, 0) <= 0 And Nz(Me.txtCostResBalance, 0) <= 0 And Nz(Me.txtAdvInvoice, 0) <= 0 Then
        strSQLField = "InvoiceExceedsTrust" 'advanced legal fee

          strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
     'INVOICE (Advanced TK) but NO ADVANCED COSTS/Fees or HOLDS and YES REPLENISH (TKEx1)
     
     ElseIf Me.txtAvailTrust > 0 And Me.txtReplBalance > 0 And Nz(Me.txtAdvancedCostBalance, 0) <= 0 And Nz(Me.txtAdvancedFeesBalance, 0) <= 0 And Nz(Me.txtCostResBalance, 0) <= 0 Then
        strSQLField = "InvoiceExceedsTrust" 'advanced legal fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
          strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        

'    ElseIf Me.txtAvailTrust > Me.txtSumofTotal And Me.txtAvailTrust < Me.txtSumofTotal + Me.txtAROutstanding And Nz(Me.RetainerReimb, 0) = 0 Then
'        strSQLField = "InvoiceExceedsTrust" 'advanced legal fee

    'INVOICE (No Advanced TK) with ADVANCED COSTS or FEES, NO HOLDS and NO REPLENISH (TKEx2)
    
    ElseIf (Nz(Me.txtAdvancedCostBalance, 0) > 0 Or Nz(Me.txtAdvancedFeesBalance, 0) > 0) And Me.txtAvailTrust > 0 And Me.txtReplBalance <= 0 And Nz(Me.txtCostResBalance, 0) <= 0 And Nz(Me.txtAdvInvoice, 0) <= 0 Then
        strSQLField = "InvoiceAdvCostFee" 'advanced legal fee
        
    'INVOICE (Advanced TK) with ADVANCED COSTS or FEES, NO HOLDS and NO REPLENISH (TKEx2)
    
    ElseIf (Nz(Me.txtAdvancedCostBalance, 0) > 0 Or Nz(Me.txtAdvancedFeesBalance, 0) > 0) And Me.txtAvailTrust > 0 And Me.txtReplBalance <= 0 And Nz(Me.txtCostResBalance, 0) <= 0 And Nz(Me.txtAdvInvoice, 0) > 0 Then
        strSQLField = "InvoiceAdvCostFee" 'advanced legal fee

         strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges


'    ElseIf Me.txtAvailTrust >= Me.txtSumofTotal + Me.txtAROutstanding And Me.RetainerReimb = -1 Then
'        strSQLField = "InvoiceNoAdvance" 'no advanced fee
'
'         strSQL = "Insert into [Matter and AR] "
'            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
'            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
'        Debug.Print strSQL
'        CurrentDb.Execute strSQL

     'INVOICE (no Advanced TK) with ADVANCED COSTS or FEES, NO HOLDS and YES REPLENISH (TKEx2)
     
     ElseIf (Nz(Me.txtAdvancedCostBalance, 0) > 0 Or Nz(Me.txtAdvancedFeesBalance, 0) > 0) And Me.txtAvailTrust > 0 And Me.txtReplBalance > 0 And Nz(Me.txtCostResBalance, 0) <= 0 And Nz(Me.txtAdvInvoice, 0) <= 0 Then
        strSQLField = "InvoiceAdvCostFee" 'reimb and advanced fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
    
    'INVOICE (Advanced TK) with ADVANCED COSTS or FEES, NO HOLDS and YES REPLENISH (TKEx2)
     
     ElseIf (Nz(Me.txtAdvancedCostBalance, 0) > 0 Or Nz(Me.txtAdvancedFeesBalance, 0) > 0) And Me.txtAvailTrust > 0 And Me.txtReplBalance > 0 And Nz(Me.txtCostResBalance, 0) <= 0 And Nz(Me.txtAdvInvoice, 0) > 0 Then
        strSQLField = "InvoiceAdvCostFee" 'reimb and advanced fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges

    'INVOICE (No Advanced TK) with any COST HOLDS and No REPLENISH (TKEx3)

    ElseIf Me.txtAvailTrust > 0 And Me.txtReplBalance <= 0 And Nz(Me.txtCostResBalance, 0) > 0 And Nz(Me.txtAdvInvoice, 0) <= 0 Then
        strSQLField = "InvoiceCostHold" 'reimb and advanced fee
        
    'INVOICE (Advanced TK) with any COST HOLDS and No REPLENISH (TKEx3)

    ElseIf Me.txtAvailTrust > 0 And Me.txtReplBalance <= 0 And Nz(Me.txtCostResBalance, 0) > 0 And Nz(Me.txtAdvInvoice, 0) > 0 Then
        strSQLField = "InvoiceCostHold" 'reimb and advanced fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges


    'INVOICE (No Advanced TK) with any COST HOLDS and YES REPLENISH (TKEx3)

     ElseIf Me.txtAvailTrust > 0 And Me.txtReplBalance > 0 And Nz(Me.txtCostResBalance, 0) > 0 And Nz(Me.txtAdvInvoice, 0) <= 0 Then
        strSQLField = "InvoiceCostHold" 'reimb and advanced fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
    'INVOICE (Advanced TK) with any COST HOLDS and YES REPLENISH (TKEx3)

     ElseIf Me.txtAvailTrust > 0 And Me.txtReplBalance > 0 And Nz(Me.txtCostResBalance, 0) > 0 And Nz(Me.txtAdvInvoice, 0) > 0 Then
        strSQLField = "InvoiceCostHold" 'reimb and advanced fee

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, AdvancedLegal, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", -1, " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges

        strSQL = "Insert into [Matter and AR] "
            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Replenish Trust Account', " & Me.txtReplBalance & ", " & maxOrderNr + 1 & ")"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
    End If


    If Me.txtReplBalance <= 0 Then
        strSQLFinal = "Update [TB Time Keeping] set " & _
                    "TrustatClose = " & Me.txtAvailTrust & ", " & _
                    "ARatClose = " & Me.txtAROutstanding & ", " & _
                    "AdvCostBal = " & Me.txtAdvancedCostBalance & ", " & _
                    "AdvFeesBal = " & Me.txtAdvancedFeesBalance & ", " & _
                    "OutsAdvDue = " & Me.txtOutstandingAdvDue & ", " & _
                    "AdvBalanceatClose = " & Me.txtCostResBalance & ", " & _
                    "ReplenishBalanceatClose = " & Me.txtReplBalance & ", " & _
                    strSQLField & "=-1, " & _
                    "[Bill Closed]=-1, " & _
                    "[Bill Closed Date] = #" & Format(Date, "yyyy-MM-dd") & "#," & _
                    "TKLocked=-1 " & _
                    " where Bill_ID=" & Me.Bill_ID
    
        CurrentDb.Execute strSQLFinal, dbSeeChanges
    
    ElseIf Me.txtReplBalance > 0 Then
        strSQLFinal = "Update [TB Time Keeping] set " & _
                    "TrustatClose = " & Me.txtAvailTrust & ", " & _
                    "ARatClose = " & Me.txtAROutstanding & ", " & _
                    "AdvCostBal = " & Me.txtAdvancedCostBalance & ", " & _
                    "AdvFeesBal = " & Me.txtAdvancedFeesBalance & ", " & _
                    "OutsAdvDue = " & Me.txtOutstandingAdvDue & ", " & _
                    "AdvBalanceatClose = " & Me.txtCostResBalance & ", " & _
                    "ReplenishBalanceatClose = " & Me.txtReplBalance & ", " & _
                    "Discount = " & Me.txtReplRequired & ", " & _
                    "TimeNotes = '*Balance Due includes requirement you maintain $" & Nz(Me.txtReplRequired, 0) & " in your trust account.', " & _
                    strSQLField & "=-1, " & _
                    "[Bill Closed]=-1, " & _
                    "[Bill Closed Date] = #" & Format(Date, "yyyy-MM-dd") & "#," & _
                    "TKLocked=-1 " & _
                    " where Bill_ID=" & Me.Bill_ID
                    
        CurrentDb.Execute strSQLFinal, dbSeeChanges
    
    End If

    Call Form_frmMatter.reorderByDateMatter(Me.CaseID)
    'Dim strBookmark As String
    'strBookmark = Me.Bookmark

    Form_frmMatter.Requery
    Me.Requery
    
    
    'Me.Bookmark = strBookmark
    
    'Fields to Update: Bill Closed, Bill Closed Date, TKLocked, Trustat Close, InvoiceTotal OR InvoiceExceedsTrust OR StatementLessTrust
    'From            : -1         , Date ()         , -1      , AvailBalance,  -1
    
    'strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
    
    '******OLD TK CLOSE CODE:
    
'    Dim strSQLField As String
'    Dim strSQLFinal As String
'    If Nz(Me.SumOfTotal, 0) = 0 Then
'        Exit Sub
'    End If
'
'    maxOrderNr = DMax("OrderNr", "[Matter and AR]", "CaseID=" & Me.CaseID)
'
'    If Me.AvailBalance >= (Me.SumOfTotal + Me.txtAROutstanding) Then
'        strSQLField = "StatementLessTrust"
'    ElseIf Me.AvailBalance < (Me.SumOfTotal + Me.txtAROutstanding) Then
'        strSQLField = "InvoiceExceedsTrust"
'
'        strSQL = "Insert into [Matter and AR] "
'            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
'            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", " & maxOrderNr + 1 & ")"
'        Debug.Print strSQL
'        CurrentDb.Execute strSQL
'
'    ElseIf Me.AvailBalance <= 0 Then
'        strSQLField = "InvoiceTotal"
'
'        strSQL = "Insert into [Matter and AR] "
'            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge, OrderNr)"
'            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Hourly Fee (" & Nz(Me.IANumber, 0) & ")', " & Me.txtAdvInvoice & ", " & maxOrderNr + 1 & ")"
'        Debug.Print strSQL
'        CurrentDb.Execute strSQL
'
'    End If
'
'    strSQLFinal = "Update [TB Time Keeping] set " & _
'                    "TrustatClose = " & Me.AvailBalance & ", " & _
'                    strSQLField & "=-1, " & _
'                    "[Bill Closed]=-1, " & _
'                    "[Bill Closed Date] = #" & Format(Date, "yyyy-MM-dd") & "#," & _
'                    "TKLocked=-1 " & _
'                    " where Bill_ID=" & Me.Bill_ID
'
'    CurrentDb.Execute strSQLFinal
'
'    Call Form_frmMatter.reorderByDateMatter(Me.CaseID)
'    'Dim strBookmark As String
'    'strBookmark = Me.Bookmark
'
'    Form_frmMatter.Requery
'    Me.Requery
'
'    'Me.Bookmark = strBookmark
'
'    'Fields to Update: Bill Closed, Bill Closed Date, TKLocked, Trustat Close, InvoiceTotal OR InvoiceExceedsTrust OR StatementLessTrust
'    'From            : -1         , Date ()         , -1      , AvailBalance,  -1
'
'    'strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID

End Sub