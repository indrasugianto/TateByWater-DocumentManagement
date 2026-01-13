' Component: Form_frmAttyFeeGeneration
' Type: document
' Lines: 38
' ============================================================

Option Compare Database

Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdFilter_Click()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
 
    If Not IsNull(Me.txtFrom) Then strSQL = strSQL & " AND tblTakeOffMonth.TakeOffDate >= #" & Me.txtFrom & "#"
    If Not IsNull(Me.txtTo) Then strSQL = strSQL & " AND tblTakeOffMonth.TakeOffDate <= #" & Me.txtTo & "#"
 
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub
Sub FilterClear()
    
    Me.txtFrom = Null
    Me.txtTo = Null
 
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdHome_Click()
     DoCmd.openform "frmhome", acNormal
End Sub