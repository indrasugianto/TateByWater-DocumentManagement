' Component: Form_frmCallsList
' Type: document
' Lines: 79
' ============================================================

Option Compare Database
Private Sub cmbAttyStaff_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdFilter_Click()
    FilterMe
End Sub

Private Sub cmdHome_Click()
DoCmd.openform "frmhome", acNormal
End Sub

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub
Private Sub cmbPracticeArea_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbReferral_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbType_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub


Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND CDate >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND CDate <= #" & Me.txtTo & "#"
    If Not IsNull(Me.cmbType) Then strSQL = strSQL & " AND ClientType = '" & Me.cmbType & "'"
    If Not IsNull(Me.txtName) Then strSQL = strSQL & " AND CallName like '*" & Me.txtName & "*'"
    If Not IsNull(Me.txtPhone) Then strSQL = strSQL & " AND CPhone like '*" & Me.txtPhone & "*'"
    If Not IsNull(Me.cmbPracticeArea) Then strSQL = strSQL & " AND CodeVal = '" & Me.cmbPracticeArea & "'"
    If Not IsNull(Me.cmbAttyStaff) Then strSQL = strSQL & " AND CAtty = '" & Me.cmbAttyStaff & "'"
    If Not IsNull(Me.txtMatter) Then strSQL = strSQL & " AND CallMatter like '*" & Me.txtMatter & "*'"
    If Not IsNull(Me.cmbReferral) Then strSQL = strSQL & " AND CReferral = '" & Me.cmbReferral & "'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub
Sub FilterClear()

    Me.txtFrom = Null
    Me.txtTo = Null
    Me.cmbType = Null
    Me.txtName = Null
    Me.txtPhone = Null
    Me.cmbPracticeArea = Null
    Me.cmbAttyStaff = Null
    Me.txtMatter = Null
    Me.cmbReferral = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub txtMatter_AfterUpdate()
    If Not IsNull(txtMatter) Then Call FilterMe
End Sub

Private Sub txtName_AfterUpdate()
    If Not IsNull(txtName) Then Call FilterMe
End Sub

Private Sub txtPhone_AfterUpdate()
    If Not IsNull(txtPhone) Then Call FilterMe
End Sub