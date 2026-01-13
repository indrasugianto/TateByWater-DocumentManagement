' Component: Form_frmClientReviews
' Type: document
' Lines: 96
' ============================================================

Option Compare Database

Private Sub CaseNo_Click()
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

Private Sub chkReviewReceived_AfterUpdate()
    Call FilterMe
End Sub

Private Sub chkSource_Click()
    Call FilterMe
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub
Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbPar_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbSource_AfterUpdate()
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
    If Not IsNull(Me.cmbPar) Then strSQL = strSQL & " AND Paralegal = '" & Me.cmbPar & "'"
    If Not IsNull(Me.cmbSource) Then strSQL = strSQL & " AND [Review Source] = '" & Me.cmbSource & "'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '" & Me.txtClient & "*'"
    If Not IsNull(Me.chkReviewReceived) Then strSQL = strSQL & " AND Reviewreceived = " & IIf(Me.chkReviewReceived, 1, 0)
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()

    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.cmbPar = Null
    Me.txtClient = Null
    Me.chkReviewReceived = Null
    Me.cmbSource = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub Form_Load()
    Call FilterMe
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub
