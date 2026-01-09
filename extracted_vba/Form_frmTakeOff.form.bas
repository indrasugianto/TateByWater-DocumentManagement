' Component: Form_frmTakeOff
' Type: document
' Lines: 666
' ============================================================

Option Compare Database

Private Sub cmdInsertIntoTA_Click()
    'GH 2017-08-11 it was not compiling so I temporary disabled it
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
    'GH 2017-08-11 it was not compiling so I temporary disabled it
    
    
    'GH 2017-08-11 it was not compiling so I temporary disabled it
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
    'GH 2017-08-11 it was not compiling so I temporary disabled it
    
    
    
    'Take off table:
    Dim strSQL2 As String
    strSQL2 = "Insert into [tblTakeOff] "
    strSQL2 = strSQL2 & "        (CaseID, TakeOffDate, EarlyEarned, TOEarned, CostReimb, CBHRev, MKRev, MTRev, CBHCom, MTCom, KBCom, MKCom, EarlyEarnedTr, TOEarnedTr, CostReimbTr)"
    
'    strSQL2 = strSQL2 & " values (" & Me.CaseID & ", #" & Me.txtTakeOffDate & "#, " & Me.txtEarlyEarned & ", " & Me.txtEarned & ", " & _
'                        Me.txtCostReimb & ", " & Me.txtCBHRev & ", " & Me.txtMKRev & ", " & Me.txtMTRev & ", " & Me.txtCBHCom & ", " & _
'                        Me.txtMTCom & ", " & Me.txtKDBCom & ", " & Me.txtMKCom & ", " & Me.ChkEarlyEarnedTR & ", " & Me.chkTOEarnedTr & ", " & Me.chkCostReimbTr & ")"
                        
                        'Fixed code behind in the Take Off form for the earlier deleted controls
                        
                        
    'GH 2017-08-11 it was not compiling so I temporary disabled it
'    strSQL2 = strSQL2 & " values (" & "" & ", #" & "" & "#, " & "" & ", " & "" & ", " & _
'                        "" & ", " & "" & ", " & "" & ", " & "" & ", " & "" & ", " & _
'                        "" & ", " & "" & ", " & "" & ", " & "" & ", " & Me.chkTOEarnedTr & ", " & Me.chkCostReimbTr & ")"
    'GH 2017-08-11 it was not compiling so I temporary disabled it
                        
    Debug.Print strSQL2
    CurrentDb.Execute strSQL2
    
    'clear fields:
    'GH 2017-08-11 it was not compiling so I temporary disabled it
    'Me.txtEarlyEarned = Null
    'Me.txtEarned = Null
    'Me.txtCostReimb = Null
    'Me.txtCBHRev = Null
    'Me.txtMKRev = Null
    'Me.txtCBHCom = Null
    'Me.txtMTCom = Null
    'Me.txtKDBCom = Null
    'Me.txtMKCom = Null
    'Me.txtMTRev = Null
    ''Me.ChkEarlyEarnedTR = Null    'deleted!
    'Me.chkTOEarnedTr = Null
    'Me.chkCostReimbTr = Null
    'GH 2017-08-11 it was not compiling so I temporary disabled it
    
    
    MsgBox "Done", , "TB CMS"

End Sub

Private Sub cmbShowReportCHHandling_Click()
    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "HandlingAtty_Case ='CMH' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
End Sub

Private Sub cmbShowReportMKHandling_Click()
    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "HandlingAtty_case ='MK' And TakeOffMonthID =" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
End Sub

Private Sub cmbShowReportMTHandling_Click()
    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "HandlingAtty_Case ='MT' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
End Sub

Private Sub cmd_PreviewReconReport_Click()
    On Error GoTo ErrHandler_cmdPreviewReconReport
    
    'GH 2017-08-11 it was not compiling so I temporary disabled it
    DoCmd.OpenReport "rptReconciliation", acViewPreview, , "TakeOffMonthID= " & Me.TakeOffMonthID
    
ErrHandler_cmdPreviewReconReport:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdAttyReport_Click()
    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "Orig_Atty='" & Me.cmbAttyTO & "' And TakeOffMonthID = " & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
    DoCmd.Close acForm, "Client_Trust_Accounts_for_Take_Off"

'"Orig_Atty ='DEB' and TakeOffMonthID=" & Me.TakeOffMonthID


End Sub




Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdRequery_Click()
    Me.Refresh
End Sub

Private Sub cmdShowReportALL_Click()
DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
On Error GoTo cmdShowReportALL_Click_Err
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", "TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
    [Report_Client Trust_Accounts_for_Take_Off].txtDate = Me.txtTakeOffDate

cmdShowReportALL_Click_Exit:
    Exit Sub

cmdShowReportALL_Click_Err:
    MsgBox Error$
    Resume cmdShowReportALL_Click_Exit
End Sub

Private Sub cmdShowReportBA_Click()
    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='BA' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
End Sub

Private Sub cmdShowReportCBH_Click()

    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "Orig_Atty ='CMH' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal



'DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
'On Error GoTo cmdShowReportCBH_Click_Err
'
'   DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='CBH'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    [Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'
'
'cmdShowReportCBH_Click_Exit:
'    Exit Sub
'
'cmdShowReportCBH_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportCBH_Click_Exit
End Sub

Private Sub cmdShowReportDEB_Click()

    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "Orig_Atty ='DEB' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal

'DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
'   On Error GoTo cmdShowReportDEB_Click_Err
'    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='DEB'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    [Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'cmdShowReportDEB_Click_Exit:
'    Exit Sub
'
'cmdShowReportDEB_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportDEB_Click_Exit
End Sub

Private Sub cmdShowReportGBF_Click()

    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "Orig_Atty ='GBF' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal



'DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
'    On Error GoTo cmdShowReportGBF_Click_Err
'    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='GBF'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    [Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'
'cmdShowReportGBF_Click_Exit:
'    Exit Sub
'
'cmdShowReportGBF_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportGBF_Click_Exit
End Sub

'------------------------------------------------------------
' Command103_Click
'
'------------------------------------------------------------
Private Sub Command103_Click()
On Error GoTo Command103_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command102" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewReport, "", "", acNormal


Command103_Click_Exit:
    Exit Sub

Command103_Click_Err:
    MsgBox Error$
    Resume Command103_Click_Exit

End Sub


'------------------------------------------------------------
' Command104_Click
'
'------------------------------------------------------------
Private Sub Command104_Click()
On Error GoTo Command104_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command103" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewReport, "", "", acNormal


Command104_Click_Exit:
    Exit Sub

Command104_Click_Err:
    MsgBox Error$
    Resume Command104_Click_Exit

End Sub


'------------------------------------------------------------
' Command105_Click
'
'------------------------------------------------------------
Private Sub Command105_Click()
On Error GoTo Command105_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command104" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewReport, "", "", acNormal


Command105_Click_Exit:
    Exit Sub

Command105_Click_Err:
    MsgBox Error$
    Resume Command105_Click_Exit

End Sub


'------------------------------------------------------------
' Command106_Click
'
'------------------------------------------------------------
Private Sub Command106_Click()
On Error GoTo Command106_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command105" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewReport, "", "", acNormal


Command106_Click_Exit:
    Exit Sub

Command106_Click_Err:
    MsgBox Error$
    Resume Command106_Click_Exit

End Sub


'------------------------------------------------------------
' Command107_Click
'
'------------------------------------------------------------
Private Sub Command107_Click()
On Error GoTo Command107_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command106" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewReport, "", "", acNormal


Command107_Click_Exit:
    Exit Sub

Command107_Click_Err:
    MsgBox Error$
    Resume Command107_Click_Exit

End Sub


'------------------------------------------------------------
' Command108_Click
'
'------------------------------------------------------------
Private Sub Command108_Click()
On Error GoTo Command108_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command107" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewReport, "", "", acNormal


Command108_Click_Exit:
    Exit Sub

Command108_Click_Err:
    MsgBox Error$
    Resume Command108_Click_Exit

End Sub


'------------------------------------------------------------
' Command109_Click
'
'------------------------------------------------------------
Private Sub Command109_Click()
On Error GoTo Command109_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command108" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewReport, "", "", acNormal


Command109_Click_Exit:
    Exit Sub

Command109_Click_Err:
    MsgBox Error$
    Resume Command109_Click_Exit

End Sub




Private Sub Command101_Click()
On Error GoTo Command101_Click_Err

    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewReport, "", "", acNormal


Command101_Click_Exit:
    Exit Sub

Command101_Click_Err:
    MsgBox Error$
    Resume Command101_Click_Exit

End Sub

Private Sub cmdShowReportJF_Click()
    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "Orig_Atty ='JF' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
End Sub

Private Sub cmdShowReportKDB_Click()

    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='KDB' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal


'DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
'   On Error GoTo cmdShowReportKDB_Click_Err
'   DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='KDB'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    [Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'
'cmdShowReportKDB_Click_Exit:
'    Exit Sub
'
'cmdShowReportKDB_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportKDB_Click_Exit
End Sub

Private Sub cmdShowReportMK_Click()

    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='MK' And TakeOffMonthID =" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal

'DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
'On Error GoTo cmdShowReportMK_Click_Err
'
'    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='MK'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    [Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'
'
'cmdShowReportMK_Click_Exit:
'    Exit Sub
'
'cmdShowReportMK_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportMK_Click_Exit
End Sub

Private Sub cmdShowReportMT_Click()

    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='MT' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal

'DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
'
'   On Error GoTo cmdShowReportMT_Click_Err
'    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='MT'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    [Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'
'cmdShowReportMT_Click_Exit:
'    Exit Sub
'
'cmdShowReportMT_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportMT_Click_Exit
End Sub

Private Sub cmdShowReportNH_Click()

    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='NH' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal

End Sub

Private Sub cmdShowReportPM_Click()

    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='PM' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal



'DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
'
'   On Error GoTo cmdShowReportPM_Click_Err
'    DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='PM'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    [Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'
'cmdShowReportPM_Click_Exit:
'    Exit Sub
'
'cmdShowReportPM_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportPM_Click_Exit
End Sub

Private Sub cmdShowReportRLF_Click()
    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='RLF' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
End Sub

Private Sub cmdShowReportTDT_Click()

    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
    strFilter = "Orig_Atty ='TDT' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal


'DoCmd.Close acReport, "Client Trust Accounts for Take Off", acSaveYes
'
'   On Error GoTo cmdShowReportTDT_Click_Err
' DoCmd.OpenReport "Client Trust Accounts for Take Off", acViewPreview, "", "Orig_Atty ='TDT'and TakeOffMonthID=" & Me.TakeOffMonthID, acNormal
'    '[Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate
'
'cmdShowReportTDT_Click_Exit:
'    Exit Sub
'
'cmdShowReportTDT_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportTDT_Click_Exit
End Sub

Private Sub cmbInsertFees_Click()

'check if reconciliation exists for this month!

'    If txtTotalEarned > 0 Then
'
'    Else
'        MsgBox "Please input an Earned fee"
'    End If
    
'    If fncRecordExists Then
'        MsgBox "You can not insert fees twice in the same month."
'        Exit Sub
'    Else
'
        'MsgBox Me.TakeOffMonthID
        
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("select * from tblTakeOffMonth where TakeOffMonthID = " & Me.TakeOffMonthID, dbOpenDynaset, dbSeeChanges)
        'rs.AddNew
        rs.Edit
                
        rs("TotalTOEarned") = Nz(Me.txtTotalEarned, 0)
        rs("TotalTOCostReimb") = Nz(Me.txtTotalCostReimb, 0)
        rs("TotalTOCommissions") = Nz(Me.txtTotalCom, 0)
        
        rs("TotalCBHRev") = Nz(Form_frmTakeOffSubForm.SumCBHRev, 0)
        rs("TotalCBHCommissions") = Nz(Form_frmTakeOffSubForm.SumCBHCom, 0)
        rs("TotalRLFCommissions") = Nz(Form_frmTakeOffSubForm.SumRLFCom, 0)
        'rs("TotalMTCommissions") = Nz(Form_frmTakeOffSubForm.SumMTCom, 0)
        rs("TotalKDBCommissions") = Nz(Form_frmTakeOffSubForm.SumKDBCom, 0)
        
        rs("JRTFees") = Nz(Me.txtJRTFees, 0)
        rs("DEBFees") = Nz(Me.TxtDEBFees, 0)
        rs("GBFFees") = Nz(Me.txtGBFFees, 0)
        rs("PMFees") = Nz(Me.txtPMFees, 0)
        rs("TDTFees") = Nz(Me.txtTDTFees, 0)
        rs("CBHFees") = Nz(Me.txtCBHFees, 0)
        rs("MKFees") = Nz(Me.txtMKFees, 0)
        'rs("MTFees") = Nz(Me.txtMTFees, 0)
        rs("KDBFees") = Nz(Me.txtKDBFees, 0)
        rs("RLFFees") = Nz(Me.txtRLFFees, 0)
        rs("NHFees") = Nz(Me.txtNHFees, 0)
        rs("JFFees") = Nz(Me.txtJFFees, 0)
        rs("WNEFees") = Nz(Me.txtWNEFees, 0)
        
        rs("FeeDataInserted") = -1
                
        rs.Update
       
        MsgBox "Done", , "TB CMS"
    'End If

    Me.cmbInsertFees.Enabled = False
    
End Sub

Function fncRecordExists() As Boolean
    strFilter = "MonthOnly = " & Month(Date) & " and YearOnly = " & Year(Date)
    ExistingID = DLookup("TakeOffMonthID", "qry_takeOff_year_month", strFilter)
    
    If ExistingID > 0 Then
        fncRecordExists = True
    Else
        fncRecordExists = False
    End If
End Function

'old code:

        'Dim strSQL As String
    
        'Insert into "Take Off Month" table:
'        strSQL = "Insert into [tblTakeOffMonth] "
'        strSQL = strSQL & "        (TakeOffDate, [WF Balance], SumUncashed, SumUncleared, WFActual, CombinedTrust, " & _
'                                   "DaleBalance, DaleActual, SomBalance, SomActual, ReconcileValue, AccReconciled, WFplusuncashed)"
'        strSQL = strSQL & " values (#" & Date & "#, " & Nz(txtSumOfBalances, 0) & ", " & Nz(txtSumOfUncashed, 0) & ", " & _
'                                    Nz(txtSumOfUnclearDeposits, 0) & ", " & Nz(txt_WF_balance_trust_amount, 0) & ", " & Nz(txtWFTrustAndJMSMTrust, 0) & ", " & _
'                                    Me.txtJMTrust & ", " & Me.txtActualJMTrust & ", " & Me.txtSMTrust & ", " & Me.txtActualSMTrust & "," & Me.txt_total_minus_inputbalance & "," & _
'                                    chkReconciled.value & ", " & Nz(Me.txt_WF_CHK, 0) & ")"

'Must insert the date in order to be able to check if rec. exists!   TakeOffDate = Date!

'        strSQL = "Insert into [tblTakeOffMonth] " & _
'                 " (TakeOffDate, TotalTOEarned, " & _
'                 "  TotalTOCostReimb, TotalTOCommissions, " & _
'                 "  TotalCBHRev, TotalCBHCommissions, " & _
'                 "  TotalMKRev, TotalMKCommissions,  " & _
'                 "  TotalMTCommissions, TotalKDBCommissions)" & _
'                 " " & _
'                 " values (" & _
'                 " " & _
'                 "#" & Date & "#, " & Nz(txtTotalEarned, 0) & ", " & _
'                 Nz(txtTotalCostReimb, 0) & ", " & Nz(txtTotalCom, 0) & ", " & Nz(txtTotalCom, 0) & ", " & _
'                 Nz(txtCBHRev, 0) & ", " & Nz(txtCBHCom, 0) & ", " & _
'                 Nz(txtMKRev, 0) & ", " & Nz(txtMKCom, 0) & ", " & _
'                 Nz(txtMTCom, 0) & ", " & Nz(txtKDBCom, 0) & ", " & _
'                 Nz(txtJRTFees, 0) & ", " & Nz(TxtDEBFees, 0) & ", " & _
'                 Nz(txtGBFFees, 0) & ", " & Nz(txtPMFees, 0) & ", " & _
'                 Nz(txtTDTFees, 0) & ", " & Nz(txtCBHFees, 0) & ", " & _
'                 Nz(txtMTFees, 0) & ", " & Nz(txtKDBFees, 0) & ", " & _
'                 Nz(txtMKFees, 0) & ")"
 
 'Debug.Print strSQL
        'CurrentDb.Execute strSQL
        
        'MaxID_tblTakeOffMonth = DMax("TakeOffMonthID", "tblTakeOffMonth")

'Private Sub Form_Current()
'    'temporary error handling:
'    On Error GoTo resumeMe
'    Me.cmbInsertFees.Enabled = Not (Me.Form.Recordset("FeeDataInserted"))
'    Exit Sub
'resumeMe:
'    Debug.Print Err.Description
'End Sub

'Private Sub Command200_Click()  NOT WORKING!
'    Form_frmTakeOffSubForm.txtTrustButton.MousePointer = 99
'    Form_frmTakeOffSubForm.txtTrustButton.MouseIcon = "d:\! Jobs\Paul Mickelsen - Access Lawyer Application\Link.cur"
'End Sub

Private Sub cmdShowReportJRT_Click()
    
    'DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    
    'On Error GoTo cmdShowReportJRT_Click_Err

    ' _AXL:<?xml version="1.0" encoding="UTF-16" standalone="no"?>
    ' <UserInterfaceMacro For="Command101" xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application" xmlns:a="http://schemas.microsoft.com/office/accessservices/2009/11/forms">
    ' _AXL:<Statements><Action Name="OpenReport"><Argument Name="ReportName">Client Trust Accounts for Take Off</Argument></Action></Statements></UserInterfaceMacro>
    'Report_Client_Trust_Accounts_for_Take_Off.txtDate.ControlSource = "=""9/1/2017""" '"=" & Me.txtTakeOffDate.value
    'Report_Client_Trust_Accounts_for_Take_Off.txtDate.Text = "=" & Me.txtTakeOffDate.value
    
    DoCmd.Close acReport, "Client_Trust_Accounts_for_Take_Off", acSaveYes
    strFilter = "Orig_Atty ='JRT' and TakeOffMonthID=" & Me.TakeOffMonthID
    DoCmd.OpenReport "Client_Trust_Accounts_for_Take_Off", acViewPreview, "", strFilter, acNormal
    
    'Report_Client_Trust_Accounts_for_Take_Off.txtDate.SetFocus
    'Report_Client_Trust_Accounts_for_Take_Off.txtDate.Text = 1111   'Me.txtTakeOffDate
    '[Report_Client Trust Accounts for Take Off].txtDate = Me.txtTakeOffDate

'cmdShowReportJRT_Click_Exit:
'    Exit Sub
'
'cmdShowReportJRT_Click_Err:
'    MsgBox Error$
'    Resume cmdShowReportJRT_Click_Exit

End Sub


