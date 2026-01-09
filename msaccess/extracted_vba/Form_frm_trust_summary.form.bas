' Component: Form_frm_trust_summary
' Type: document
' Lines: 170
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

Private Sub chkNonZero_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbAddressLabel_Click()
On Error GoTo ErrHandler_cmdPrintLabelP_Click
    
    If IsNull(Me.Executor) Then
    
    answer = MsgBox("Would you like to print this label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_address_label").IsLoaded Then
            DoCmd.Close acReport, "rpt_address_label", acSaveNo
        End If
        
        DoCmd.OpenReport "rpt_address_label", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    End If
   
    If Not IsNull(Me.Executor) Then
    
    answer = MsgBox("Would you like to print this label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_address_labelEx").IsLoaded Then
            DoCmd.Close acReport, "rpt_address_labelEx", acSaveNo
        End If
        
        DoCmd.OpenReport "rpt_address_labelEx", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    End If
ErrHandler_cmdPrintLabelP_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub


Private Sub cmbHome_Click()
    
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbPracticeArea_Click()
    Call FilterMe
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub

Private Sub cmdPreviewTrustStmt_Click()
 
On Error GoTo ErrHandler_PreviewTrustStmt_Click
    
    If CurrentProject.AllReports("Statement of Trust Account").IsLoaded Then
        DoCmd.Close acReport, "Statement of Trust Account", acSaveYes
    End If
    StatementofTrustAccount_Filter = True
    DoCmd.OpenReport "Statement of Trust Account", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
    
ErrHandler_PreviewTrustStmt_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub



' On Error GoTo ErrHanlder_cmdStatementrustAccount
'        StatementofTrustAccount_Filter = True
'        DoCmd.OpenReport "Statement of Trust Account", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
'ErrHanlder_cmdStatementrustAccount:
'    If Err.Number = 2501 Then
'        ShowMessage "No records found!"
'    ElseIf Err.Number <> 0 Then
'        ShowMessage Err.Description
'    End If
'End Sub

Private Sub cmdPrintStmtTrust_Click()
On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("Statement of Trust Account").IsLoaded Then
            DoCmd.Close acReport, "Statement of Trust Account", acSaveNo
        End If
        
        DoCmd.OpenReport "Statement of Trust Account", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '" & Me.txtClient & "*'"
    If Not IsNull(Me.cmbPracticeArea) Then strSQL = strSQL & " AND Case_Letter like '" & Me.cmbPracticeArea & "*'"
    
    If Not IsNull(Me.chkNonZero) Then
        If chkNonZero Then
            strSQL = strSQL & " AND Balance >0 "
        Else
            strSQL = strSQL & " AND Balance =0 "
        End If
    End If
    
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
    Me.chkNonZero = Null
    Me.cmbOrigAtty = Null
    Me.txtClient = Null
    Me.cmbPracticeArea = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub Combo13_Click()
    
End Sub

Private Sub Form_Load()
'    Me.filter = "Balance > 0 or Balance < 0"
'    Me.FilterOn = True
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub