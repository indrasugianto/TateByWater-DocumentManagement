' Component: Form_frmHome
' Type: document
' Lines: 292
' ============================================================


Private Sub cmbAdvancedExp_Click()
    DoCmd.openform "frm_advanced_payments"
End Sub

Private Sub cmbScanForm_Click()
    DoCmd.openform "frmToBeScanned"
End Sub

Private Sub cmbTimeKeepingOpen_Click()
    DoCmd.openform "frmTimeKeepingOpen"
End Sub

Private Sub cmbTo_BeCloseForm_Click()
    DoCmd.openform "frmToBeClosed"
End Sub

Private Sub cmdActionNeeded_Click()
    DoCmd.openform "frmActionNeededAll"
End Sub

Private Sub cmdBillingTracker_Click()
    DoCmd.openform "frm_Billing_Tracker2"
End Sub

Private Sub cmdCallsForm_Click()
    DoCmd.openform "frmCalls"
End Sub

Private Sub cmdCaseList_Click()
    DoCmd.openform "frmCaseList", acNormal, , , , , Right(Year(Date), 2)
    'Util.openform "frmYearWiseCaseList"
End Sub

Private Sub cmdConflictChk_Click()
    DoCmd.openform "frmConflictChk"
End Sub

Private Sub cmdDispositions_Click()
    DoCmd.openform "frmDispositions"
End Sub

Private Sub cmdPIStatus_Click()
    DoCmd.openform "frmPersInjuryStatusReport"
End Sub

Private Sub cmdReceipt_Click()
    DoCmd.openform "frmReceipt"
End Sub

Private Sub cmdSourceAnalytics_Click()
    DoCmd.openform "frmSourceAnalytics"
End Sub

Private Sub cmdTakeOffSteps_Click()
    DoCmd.openform "frmTakeOffSteps"
End Sub

Private Sub cmdTKBilling_Click()
    DoCmd.openform "frmTimeKeepingClosed"
End Sub

Private Sub cmdTrust_Click()
    DoCmd.openform "frm_trust_summary"
End Sub

Private Sub cmdTrustAccountChron_Click()
    DoCmd.openform "frmTrustEntriesChron"
End Sub

Private Sub cmdUpcomingHrgs_Click()
 DoCmd.openform "frmUpcoming Hearings"
End Sub

Private Sub Command10_Click()
'    IsDisableEvents = False
'    OpenFormHomeScreen "frmClientLedger"
    DoCmd.openform "frmClientLedger"
    DoCmd.ApplyFilter "[Yr]= '" & Right(Year(Now), 2) & "' and Closed=0"
    DoCmd.GoToRecord , , acNewRec
'    Me.filter = "[Yr]= '" & Right(Year(Now), 2) & "' and Closed=0"
'    Me.FilterOn = True
'    Add New record:
'    DoCmd.GoToRecord , , acNewRec
'    DoCmd.openform formName:="frmHome", _
    WhereCondition:="[yr] = " & Right(Year(Now), 2) & "'"
End Sub

Private Sub Command102_Click()
    
End Sub

Private Sub Command11_Click()
    OpenFormHomeScreen "Intakes"
End Sub

Private Sub cmdInvoice_Click()
    On Error GoTo ErHandler_CmdInvoice
    
    DoCmd.openform "frm_invoices_summary"
    'DoCmd.OpenReport "Invoice", acViewPreview
    
ErHandler_CmdInvoice:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdDeleteAll_Click()
'
'    answer = MsgBox("Are you 100% sure you want to delete all the data in the database?", vbYesNo, "TB CMS: Delete all data")
'
'    If answer = vbYes Then
'        For Each t In CurrentDb.TableDefs
'            'Debug.Print t.Name
'            If Not InStr(LCase(t.Name), "msys") > 0 Then
'                'Debug.Print t.Name
'                strSQL = "delete * from [" & t.Name & "]"
'                CurrentDb.Execute strSQL
'            End If
'        Next
'    End If
    
End Sub

'Private Sub cmbCloseForm_Click()
'    DoCmd.openform "frm_trust_summary"
'End Sub

'Private Sub cmbPISOL_Click()
'    DoCmd.openform "frmPersInjSOL"
'End Sub

'Private Sub cmdInvoiceAttachDiscount_Click()
'    On Error GoTo ErHandler_cmdInvoiceAttachDiscount
'    DoCmd.openform "frmSelectCaseNum_Discount"
'ErHandler_cmdInvoiceAttachDiscount:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Private Sub cmdInvoiceAttachHourly_Click()
'    On Error GoTo ErHandler_cmdInvoiceAttachHourly
'    DoCmd.OpenReport "Invoice Attach - Hourly", acViewPreview
'ErHandler_cmdInvoiceAttachHourly:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Private Sub cmdInvoiceNoBalanceDue_Click()
'     On Error GoTo ErHandler_cmdInvoiceNoBalanceDue
'    DoCmd.OpenReport "Invoice - No Balance Due", acViewPreview, , "ChkBalanceDue=" & 0
'ErHandler_cmdInvoiceNoBalanceDue:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Private Sub cmdStatementofTrust_Click()
'    On Error GoTo ErHandler_cmdStatementofTrust
'
'    DoCmd.OpenReport "Statement of Trust Account", acViewPreview
'ErHandler_cmdStatementofTrust:
'    If Err.Number = 2501 Then
'        ShowMessage "No Record Found!"
'     ElseIf Err.Number <> 0 Then ShowMessage Err.Description
'    End If
'End Sub

'Private Sub cmdTrustEntriesChronological_Click()
'    On Error GoTo ErHandler_cmdTrustEntriesChronological
'    DoCmd.openform "frmTrustEntriesChron"
'ErHandler_cmdTrustEntriesChronological:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Private Sub Command13_Click()
'    lngChnc = 2
'    OpenFormHomeScreen "frmFamilyLaw"
'End Sub

'Private Sub Command19_Click()
'     'docmd.OpenReport "Invoice Attach - Hourly (12 entries)",acViewPreview,
'    OpenFormHomeScreen "frmSelectCaseNum"
'End Sub

'Private Sub Command21_LostFocus()
''    If Me.Command12.Visible = True Then
''        Me.Command12.Visible = False
''        Me.Command19.Visible = False
''    End If
'End Sub

'Private Sub Command26_Click()
'    On Error GoTo ErrHandler_Command26
'        DoCmd.OpenReport "Accounts Receivable", acViewPreview
'ErrHandler_Command26:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Private Sub Command36_Click()
'    On Error GoTo ErrHandler_Command36
'        DoCmd.OpenReport "rpt_Main_Closing", acViewPreview
'ErrHandler_Command36:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Private Sub Image17_Click()
'    If MsgBox("Do you want to exit the application?", vbYesNo + vbCritical) = vbYes Then DoCmd.Quit acQuitSaveNone
'End Sub

'Private Sub cmdInvoicePastDue_Click()
'    'On Error Resume Next
'    DoCmd.Close acReport, "Invoice - Past Due", acSaveYes
'
'    On Error GoTo ErHandler_cmdInvoicePastDue
'    DoCmd.OpenReport "Invoice - Past Due", acViewPreview, , "[Past Due]=" & -1
'ErHandler_cmdInvoicePastDue:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Gaz Removed:

'Private Sub Command21_Click()
'    Me.Command12.Visible = Not Me.Command12.Visible
'    Me.Command19.Visible = Not Me.Command19.Visible
'End Sub

'Private Sub Command12_Click()
'    OpenFormHomeScreen "Time Keeping"
'End Sub

'Private Sub cmdAdvancedPayments_Click()
'    On Error GoTo ErHandler_cmdAdvancedPayments_Click
'
'    DoCmd.openform "frm_advanced_payments"
'
'ErHandler_cmdAdvancedPayments_Click:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub

'Private Sub cmdUncashed_Click()
'    On Error GoTo ErHandler_cmdUncashed_Click
'
'    DoCmd.openform "frm_uncashed_trust_checks"
'
'ErHandler_cmdUncashed_Click:
'    If Err.Number <> 0 Then ShowMessage Err.Description
'End Sub


Private Sub Command100_Click()
    'DoCmd.ShowToolbar "Ribbon", acToolbarNo
    EnableProperties
End Sub

Private Sub Command101_Click()
    DisableProperties
End Sub

Private Sub Command115_Click()
    DoCmd.openform "frmClientReviews"
End Sub

Private Sub Command117_Click()
    DoCmd.openform "frmCalendarCheck"
End Sub

Private Sub Command119_Click()
    DoCmd.openform "frmDispositions"
End Sub

Private Sub Command120_Click()
  DoCmd.openform "frmHomeAdminLogin"
  
'  answer = InputBox("Please input the Admin Pasword", "TB CMS: Admin Pass Required")
'    If answer = "27admin40" Then
'    DoCmd.openform "frmHomeAdmin"
'    Else
'        Me.Undo
'    End If
End Sub

Private Sub Cr_Report_Click()
    DoCmd.openform "frmCrimStatusReport"
End Sub

Private Sub Form_Load()
    If isACCDE Then
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If
    
'    If frmLogin.UserID.Text = "admin" Then
'        cmdTakeOffSteps.Visible = True
'    Else
'        cmdTakeOffSteps.Visible = False
'    End If
End Sub