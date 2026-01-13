' Component: Form_frmActionNeededAll
' Type: document
' Lines: 180
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

Private Sub cmbClearFilter_Click()

End Sub

Private Sub cmbAttyStaff_AfterUpdate()
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

Private Sub cmbPar_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdActionNeededDone_Click()
On Error GoTo ERR_HANDLER:
    Dim lActionNeededID As Long
    Dim cn As ADODB.Connection
    Dim sql As String
    
    lActionNeededID = pcaConvertNulls(Me.ActionNeededID, 0)
    
    If lActionNeededID <> 0 Then
        Set cn = New ADODB.Connection
        cn.Open PcaGetConnnectionString()
        
        sql = ""
        sql = sql & "UPDATE TblActionNeeded "
        sql = sql & "SET ActionComp = CASE WHEN ActionComp = 1 THEN 0 ELSE 1 END "
        sql = sql & "WHERE ActionNeededID = " & lActionNeededID
        
        cn.Execute sql
        
        Me.Refresh
        
    
    End If
EXIT_HANDLER:
    Exit Sub
ERR_HANDLER:
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
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
    If Not IsNull(Me.cmbCodeVal) Then strSQL = strSQL & " AND CodeVal = '" & Me.cmbCodeVal & "'"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.cmbPar) Then strSQL = strSQL & " AND Paralegal = '" & Me.cmbPar & "'"
    If Not IsNull(Me.txtMatter) Then strSQL = strSQL & " AND Matter_type like '*" & Me.txtMatter & "*'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '*" & Me.txtClient & "*'"
    If Not IsNull(Me.cmbAttyStaff) Then strSQL = strSQL & " AND ActPerson = '" & Me.cmbAttyStaff & "'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()

    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.cmbPar = Null
    Me.cmbCodeVal = Null
    Me.txtMatter = Null
    Me.txtClient = Null
    Me.cmbAttyStaff = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub


Private Sub DateComp1_DblClick(Cancel As Integer)
    Dim userInput As String
    Dim userDate As Date
    Dim validDate As Boolean

    Dim lActionNeededID As Long
    Dim cn As ADODB.Connection
    Dim sql As String
    
    lActionNeededID = pcaConvertNulls(Me.ActionNeededID, 0)
    
    If lActionNeededID <> 0 Then
        ' Loop until the user enters a valid date or cancels the input
        Do
            ' Prompt the user to enter a date
            userInput = InputBox("Please enter a date (MM/DD/YYYY):", "Enter Date")
    
            ' If the user clicks Cancel or leaves the input blank, exit the loop
            If userInput = "" Then
                MsgBox "No date entered"
                Exit Sub
            End If
    
            ' Validate the entered date
            On Error Resume Next
            userDate = CDate(userInput)
            validDate = Not (Err.Number <> 0)
            On Error GoTo 0
    
            ' If valid, exit the loop
            If validDate Then
                Set cn = New ADODB.Connection
                cn.Open PcaGetConnnectionString()
            
                sql = ""
                sql = sql & "UPDATE TblActionNeeded "
                sql = sql & "SET DateComp1 = " & pcaAddQuotes(userDate)
                sql = sql & " WHERE ActionNeededID = " & lActionNeededID
                
                cn.Execute sql
                Me.Refresh
                
            Else
                MsgBox "Invalid date. Please try again."
            End If
        Loop Until validDate
    End If

End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub

Private Sub txtMatter_AfterUpdate()
    Call FilterMe
End Sub