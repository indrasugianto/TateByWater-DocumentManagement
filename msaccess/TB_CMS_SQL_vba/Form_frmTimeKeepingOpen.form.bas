' Component: Form_frmTimeKeepingOpen
' Type: document
' Lines: 225
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

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '" & Me.txtClient & "*'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub ChkCloseTKFilter_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbclose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmbHome_Click()
DoCmd.openform "frmhome", acNormal
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Sub FilterClear()
    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.txtClient = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo ErrHandler_CaseID_Click
    
'    If Me.Discount > 0 Then
'        strReportName = "Invoice Attach - Hourly w Discount"
'    Else
'        strReportName = "Invoice Attach - Hourly"
'    End If
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
    
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
    
    'GH 2017-08-11 it was not compiling so I temporary disabled it
    'DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
    
ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdPrintInvoice_Click()
    On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
'    If Me.Discount > 0 Then
'        strReportName = "Invoice Attach - Hourly w Discount"
'    Else
'        strReportName = "Invoice Attach - Hourly"
'    End If
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports(strReportName).IsLoaded Then
            DoCmd.Close acReport, strReportName, acSaveNo
        End If
        
        'GH 2017-08-11 it was not compiling so I temporary disabled it
       DoCmd.OpenReport strReportName, acNormal, , "[Bill_ID]=" & Me.Bill_ID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description

End Sub

Private Sub IANumber_Click()
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
    
    Form_frmClientLedger.tabControl.Pages(4).SetFocus
    [Form_Time Keeping].cmbBills = Me.txtBillID
    [Form_Time Keeping].cmbBills_AfterUpdate
    
ErrHandler_CaseNum_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdAddNewTK_Click()
    If CurrentProject.AllForms("frmClientLedger").IsLoaded Then
        DoCmd.Close acForm, "frmClientLedger", acSaveNo
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbCases, , , Me.cmbCases
        Forms("frmClientLedger").SetFocus
    Else
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbCases, , , Me.cmbCases
        Forms("frmClientLedger").SetFocus
        End If
    Form_frmClientLedger.tabControl.Pages(4).SetFocus

'
'    If Me.Dirty Then Me.Dirty = False
'    'check if this is the first TK.
'    strCaseID = Form_frmClientLedger.CaseID
'    TKcount = DCount("Bill_ID", "[TB Time Keeping]", "CaseID=" & strCaseID)
'    If TKcount > 0 Then
'        'if this is not the first TK then check if all previous TKs are closed
'        TKcountClosed = DCount("Bill_ID", "[TB Time Keeping]", "CaseID=" & strCaseID & " and [Bill Closed]=-1")
'        If TKcount = TKcountClosed Then
'            'all TKs are closed, Good to go.
'            'Debug.Print "OK"
'            If CurrentProject.AllForms("frmClientLedger").IsLoaded Then
'            DoCmd.Close acForm, "frmClientLedger", acSaveNo
'            DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbCases
'            Forms("frmClientLedger").SetFocus
'            Else
'            DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbCases
'            Forms("frmClientLedger").SetFocus
'            End If
'            Form_frmClientLedger.tabControl.Pages(4).SetFocus
'            Call [Form_Time Keeping].cmbBills_AfterUpdate
'            [Form_Time Keeping].Requery
'            Call [Form_Time Keeping].addNewTK(strCaseID)
'        Else
'            'error - can't go further.
'            'Debug.Print "Error"
'            MsgBox "You cannot add a new TK until the current TK is closed." & vbCrLf & "You can find open TK below.", vbInformation, "TB CMS"
'        End If
'    Else
''    If Not IsNull(strCaseID) Then
''        currNr = Nz(DLookup("CountOfIANumber", "qry_get_time_keeping_numbers", "CaseID = " & strCaseID), 0)
''        strIANumber = "TK-" & Nz(currNr) + 1
''    End If
''
''    strSQL = "Insert into [TB Time Keeping] (CaseID, [Bill Open], Discount, IANumber)"
''    strSQL = strSQL & " values (" & strCaseID & ", #" & Date & "#, 0, '" & strIANumber & "')"
''
''    Debug.Print "strSQL=" & strSQL
''    CurrentDb.Execute strSQL
''
''    Debug.Print "strIANumber=" & strIANumber
''
''    [Form_Time Keeping].cmbBills.RowSource = "select * from qryBillList where CaseID=" & strCaseID
''    [Form_Time Keeping].cmbBills.Requery
''
''    maxBillID = DMax("Bill_ID", "qryBillList", "CaseID=" & strCaseID)
''
'
'    '[Form_Time Keeping].Recordset.MoveLast
'  End If
End Sub



'Private Sub cmdRecord_Click()
'    strSQL = "Insert into [tbl_InvoiceSent] "
'            strSQL = strSQL & "        (CaseID, TKDate, [TK Sent], TKNumber)"
'            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL
'
'
'    strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL
'
'    'Me.[Bill Sent].Requery
'    Me.Requery
'    MsgBox "Recorded", , "TB CMS"
'End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub