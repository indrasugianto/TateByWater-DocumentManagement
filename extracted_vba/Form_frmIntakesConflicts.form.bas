' Component: Form_frmIntakesConflicts
' Type: document
' Lines: 75
' ============================================================

Option Compare Database

Private Sub cmbClearFilter_Click()
    Call FilterClear
End Sub



Private Sub Text0_Click()

End Sub

Private Sub IDNum_Click()
On Error GoTo ErrHandler_IDNum_Click
    IsDisableEvents = True
    If CurrentProject.AllForms("Intakes").IsLoaded Then
        DoCmd.Close acForm, "Intakes", acSaveNo
        DoCmd.openform "Intakes", acNormal, , "[ID]=" & Me.ID
'        Forms("Intakes").SetFocus
    Else
        DoCmd.openform "Intakes", acNormal, , "[ID]=" & Me.ID
'        Forms("Intakes").SetFocus
    End If
ErrHandler_IDNum_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub





Private Sub txtDOBGI_AfterUpdate()
If Not IsNull(txtDOBGI) Then Call FilterMe
End Sub

Private Sub txtFirstGI_AfterUpdate()
 If Not IsNull(txtFirstGI) Then Call FilterMe
End Sub

Private Sub txtLastGI_AfterUpdate()
 If Not IsNull(txtLastGI) Then Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.txtLastGI) Then strSQL = strSQL & " AND [GI Last Name] like '*" & Me.txtLastGI & "*'"
    If Not IsNull(Me.txtFirstGI) Then strSQL = strSQL & " AND [GI First Name] like '*" & Me.txtFirstGI & "*'"
    If Not IsNull(Me.txtDOBGI) Then strSQL = strSQL & " AND GIDOB like '" & Me.txtDOBGI & "*'"
    If Not IsNull(Me.txtMatterGI) Then strSQL = strSQL & " AND [GI Matter] like '*" & Me.txtMatterGI & "*'"

    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()
    'clear controls:
    Me.txtLastGI = Null
    Me.txtFirstGI = Null
    Me.txtDOBGI = Null
    Me.txtMatterGI = Null
'    Me.chkPastDue = Null
'    Me.chkNoBalance = Null
'    Me.chkNonZero = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub txtMatterGI_AfterUpdate()
    If Not IsNull(txtMatterGI) Then Call FilterMe
End Sub