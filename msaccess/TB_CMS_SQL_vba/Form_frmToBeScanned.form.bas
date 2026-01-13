' Component: Form_frmToBeScanned
' Type: document
' Lines: 77
' ============================================================

Option Compare Database
Private cls As New clsFormValidation

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
    If Not IsNull(Me.cmbYear) Then strSQL = strSQL & " AND YR like '" & Me.cmbYear & "*'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND ClientName like '" & Me.txtClient & "*'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
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
    Me.cmbYear = Null
    Me.txtClient = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub


Private Sub cmbYear_AfterUpdate()
   Call FilterMe
End Sub


Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub