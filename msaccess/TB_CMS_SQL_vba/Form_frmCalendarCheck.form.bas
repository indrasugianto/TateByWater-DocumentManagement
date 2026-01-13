' Component: Form_frmCalendarCheck
' Type: document
' Lines: 94
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

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbHrgType_AfterUpdate()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbPar_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub cmdFilter_Click()
    Call FilterMe
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub

Private Sub txtMatter_AfterUpdate()
    If Not IsNull(txtMatter) Then Call FilterMe
End Sub
Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbHrgType) Then strSQL = strSQL & " AND HearingType = '" & Me.cmbHrgType & "'"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.cmbPar) Then strSQL = strSQL & " AND Paralegal = '" & Me.cmbPar & "'"
    If Not IsNull(Me.txtMatter) Then strSQL = strSQL & " AND Matter_type like '*" & Me.txtMatter & "*'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '*" & Me.txtClient & "*'"
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND Hearing_Date >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND Hearing_Date <= #" & Me.txtTo & "#"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()

    Me.cmbClients = Null
    Me.cmbHrgType = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.cmbPar = Null
    Me.txtMatter = Null
    Me.txtFrom = Null
    Me.txtTo = Null
    Me.txtClient = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub cmdHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub