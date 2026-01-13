' Component: Form_frmYearWiseCaseList
' Type: document
' Lines: 103
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
    
    If Not IsNull(Me.chkClosed) Then strSQL = strSQL & " AND Closed = " & Me.chkClosed
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbCodeVal) Then strSQL = strSQL & " AND CodeVal = '" & Me.cmbCodeVal & "'"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.txtMatter) Then strSQL = strSQL & " AND Matter_type like '" & Me.txtMatter & "*'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub
Private Sub chkClosed_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub


Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbCodeval_AfterUpdate()
    Call FilterMe
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

Private Sub Combo25_AfterUpdate()
    Me.FilterOn = False
    Me.Filter = "[YR]='" & Me.Combo25 & "'"
    Me.FilterOn = True
End Sub

Private Sub Command225_Click()
cls.ExeCommand SaveRec
End Sub

Private Sub Command227_Click()
'    cls.ExeCommand Cancelrec
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.Combo25.value = Me.OpenArgs
        Combo25_AfterUpdate
    End If
    Set cls.Form = Me
End Sub

Private Sub txtMatter_AfterUpdate()
    If Not IsNull(txtMatter) Then Call FilterMe
End Sub

Sub FilterClear()
   
    Me.cmbClients = Null
    Me.cmbCodeVal = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.txtMatter = Null
    Me.chkClosed = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub
