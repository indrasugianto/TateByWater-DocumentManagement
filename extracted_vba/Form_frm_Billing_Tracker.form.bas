' Component: Form_frm_Billing_Tracker
' Type: document
' Lines: 65
' ============================================================

Option Compare Database

Private Sub cmbClearFilter_Click()
    Call FilterClear
End Sub

'Private Sub Case_Letter_AfterUpdate()
'    Call FilterMe
'End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND Tdate >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND Tdate <= #" & Me.txtTo & "#"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Tatty = '" & Me.cmbOrigAtty & "'"
        
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub cmbclose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmbCodeVal_Click()
    Call FilterMe
End Sub

Sub FilterClear()

    Me.txtFrom = Null
    Me.txtTo = Null
    Me.cmbOrigAtty = Null

    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmbFilter_Click()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub cmbHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub