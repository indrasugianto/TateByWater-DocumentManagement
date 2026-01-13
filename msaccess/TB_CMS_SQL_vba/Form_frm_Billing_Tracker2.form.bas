' Component: Form_frm_Billing_Tracker2
' Type: document
' Lines: 164
' ============================================================

Option Compare Database
Option Explicit

Private Sub cmbClearFilter_Click()
    Call FilterClear
End Sub

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
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbTAtty) Then strSQL = strSQL & " AND Tatty = '" & Me.cmbTAtty & "'"
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND Name = '" & Me.cmbClients & "'"
        
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
    
    Me.cmbTAtty = Null
    Me.Tatty = Null
    
    Me.cmbClients = Null

    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmbFilter_Click()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbTAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdAttyReport_Click()
    If pcaempty(txtFrom) Or pcaempty(txtTo) Then
        MsgBox "Please specify the date range for the report"
    Else
        DoCmd.OpenReport "rptBillingTotals", acViewPreview
    End If
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
Private Function GetTotalHours() As Long
On Error GoTo ERR_HANDLER
Dim TotalHour As Long
Dim sql As String
Dim swhere As String
Dim rs As Recordset
   
    swhere = Me.Filter
    
    
    
    sql = ""
    sql = sql & " SELECT SUM([Time_]) as TotalHours FROM qryBillingTracker2 "
    If Len(swhere) > 2 Then
        sql = sql & " WHERE " & swhere
    End If
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        TotalHour = rs("TotalHours")
    Else
        TotalHour = 0
    End If
    
EXIT_HANDLER:
    GetTotalHours = TotalHour
    Exit Function
ERR_HANDLER:
    TotalHour = 0
    Resume EXIT_HANDLER
End Function

Private Function GetTotalBilled() As Currency
On Error GoTo ERR_HANDLER
Dim TotalBilled As Currency
Dim sql As String
Dim swhere As String
Dim rs As Recordset
   
    swhere = Me.Filter
       
    
    sql = ""
    sql = sql & " SELECT SUM([Billed]) as TotalBilled FROM qryBillingTracker2 "
    If Len(swhere) > 2 Then
        sql = sql & " WHERE " & swhere
    End If
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then
        TotalBilled = rs("TotalBilled")
    Else
        TotalBilled = 0
    End If
EXIT_HANDLER:
    GetTotalBilled = TotalBilled
    Exit Function
ERR_HANDLER:
    TotalBilled = 0
    Resume EXIT_HANDLER
End Function
