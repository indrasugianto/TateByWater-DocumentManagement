' Component: Form_frmToBeClosed
' Type: document
' Lines: 110
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
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND ClientName like '" & Me.txtClient & "*'"
'    If Not IsNull(Me.chkZero) Then
'        If chkZero Then
'            strSQL = strSQL & " AND [Current balance].value = 0 and txtTrustBalance.value = 0 "
'        Else
'            strSQL = strSQL & " AND [Current balance].value <> 0 and txtTrustBalance.value <> 0 "
'        End If
'    End If
    
    
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub

'Private Sub chkZero_AfterUpdate()
'    Call FilterMe
'End Sub

Private Sub Closed_BeforeUpdate(Cancel As Integer)
    If [Current Balance] <> 0 Or txtTrustBalance <> 0 Then
        MsgBox "Client AR Balance and Trust Account Balance Must Both Be 0 to Close.", vbExclamation, "TB CMS"
        Me.Undo
    End If
End Sub
Private Sub CloseDate_BeforeUpdate(Cancel As Integer)
    If Me.Closed = False Then
        MsgBox "Please check Closed Box Before Entering Date.", vbExclamation, "TB CMS"
        Me.Undo
    End If
End Sub

Private Sub cmbClearFilter_Click()
      Call FilterClear
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

Sub FilterClear()
   
    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.txtClient = Null
'    Me.chkZero = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub


Private Sub Combo25_AfterUpdate()
    Me.FilterOn = False
    Me.Filter = "[YR]='" & Me.Combo25 & "'"
    Me.FilterOn = True
End Sub

Private Sub Command353_Click()
  On Error GoTo ErrHandler_CmdClosingSheet
    DoCmd.Close acReport, "rpt_Main_Closing", acSaveYes
    DoCmd.OpenReport "rpt_Main_Closing", acViewPreview, , "[CaseID]=" & Me.CaseID
ErrHandler_CmdClosingSheet:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub