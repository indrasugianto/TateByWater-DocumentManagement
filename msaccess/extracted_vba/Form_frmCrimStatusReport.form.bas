' Component: Form_frmCrimStatusReport
' Type: document
' Lines: 125
' ============================================================

Option Compare Database

Private Sub cmdAllAttyReport_Click()
    DoCmd.Close acReport, "rptCriminalStatus", acSaveYes
    DoCmd.OpenReport "rptCriminalStatus", acViewPreview
    DoCmd.Close acForm, "frmCrimStatusReport"
End Sub

Private Sub cmdAttyReport_Click()
    DoCmd.Close acReport, "rptCriminalStatus", acSaveYes
    'DoCmd.OpenReport "rptCriminalStatus", acViewPreview, "", "Orig_Atty='" & Me.cmbCrAtty & "'"
    
    Dim strWhere As String
    
    If Len(Me.cmbCrAtty & "") > 0 Then
    strWhere = "[Orig_Atty]='" & Me.cmbCrAtty & "' "
    End If
    
    If Len(Me.cmbCrCourt & "") > 0 Then
    strWhere = "[Court]='" & Me.cmbCrCourt & "' "
    End If
    
    If Len(Me.txtMatter & "") > 0 Then
    strWhere = "[Matter_Type]LIKE """ & Me.txtMatter & "*"" "
    End If
    
    If strWhere <> "" Then
    If Len(Me.txtMatter & "") <> 0 Then
    If Len(Me.cmbCrCourt & "") <> 0 Then
    If Len(Me.cmbCrAtty & "") <> 0 Then
    strWhere = "[Orig_Atty] ='" & Me.cmbCrAtty & "' And [Court]='" & Me.cmbCrCourt & "' And [Matter_Type] LIKE """ & Me.txtMatter & "*"""
    Else
    strWhere = "[Court]='" & Me.cmbCrCourt & "'"
    End If
    End If
    End If
    End If
    DoCmd.OpenReport "rptCriminalStatus", acViewPreview, , strWhere
   
    DoCmd.Close acForm, "frmCrimStatusReport"
    
'   WORKS FOR TWO!
'Dim strWhere As String
'
'    If Len(Me.cmbCrAtty & "") > 0 Then
'    strWhere = "[Orig_Atty]='" & Me.cmbCrAtty & "' "
'    End If
'
'    If Len(Me.cmbCrCourt & "") > 0 Then
'    strWhere = "[Court]='" & Me.cmbCrCourt & "' "
'    End If
'
'    If strWhere <> "" Then
'    If Len(Me.cmbCrCourt & "") <> 0 Then
'    If Len(Me.cmbCrAtty & "") <> 0 Then
'    strWhere = "[Orig_Atty] ='" & Me.cmbCrAtty & "' And [Court]='" & Me.cmbCrCourt & "'"
'    Else
'    strWhere = "[Court]='" & Me.cmbCrCourt & "'"
'    End If
'    End If
'    End If
'    DoCmd.OpenReport "rptCriminalStatus", acViewPreview, , strWhere
'
'    DoCmd.Close acForm, "frmCrimStatusReport"
    
'    whereStmt = ""
'    If Not IsNull(Me.cmbCrAtty) Then
'    whereStmt = "Orig_Atty= LIKE """ & Me.cmbCrAtty & "*"" AND "
'    End If
'    If Not IsNull(Me.cmbCrCourt) Then
'    whereStmt = whereStmt & "Court LIKE """ & Me.cmbCrCourt & "*"" AND "
'    End If
'    If Not IsNull(Me.txtMatter) Then
'    whereStmt = whereStmt & "[Matter_Type] LIKE """ & Me.txtMatter & "*"" AND "
'    End If
'    whereStmt = Left(whereStmt, Len(whereStmt) - 4)
'    DoCmd.OpenReport "rptCriminalStatus", acViewPreview, whereStmt
'
'    DoCmd.Close acForm, "frmCrimStatusReport"
    'DoCmd.OpenReport "rptCriminalStatus", acViewPreview, "", "Orig_Atty='" & Nz(Me.cmbCrAtty) & "' AND Court='" & Nz(Me.cmbCrCourt) & "' AND Matter_Type = '" & Nz(Me.txtMatter) & "'"
    
    
'   On Error GoTo Error_Handler
'

    
    
    
    
    
    
    'docmd.openreport "MidModReport", acViewReport,, "Mod = '" & cboMod & "' AND studentID = '" cboStudent & "'".
    'DoCmd.OpenReport "MidModReport", acViewReport,, "Mod = '" & cboMod & "' And studentID = '" & cboStudent & "'"
    'AND Matter_Type LIKE '" & txtMatter & "'"
    
'Dim whereStmt As String

'whereStmt = ""

'If Not IsNull(Me.txtEstimator) Then
'whereStmt = "[Estimator] = " & Me.txtEstimator & " AND "
'End If
'If Not IsNull(Me.txtYear) Then
'whereStmt = whereStmt & "[SalesYear] = " & Me.txtYear & " AND "
'End If

'whereStmt = Left(whereStmt, Len(whereStmt) - 5)
    
    
    
':
'WHEREStmt = ""
'If Not IsNull(Me.txtEstimator) Then
'WHEREStmt = "[Estimator] = " & Me.txtEstimator & " AND "
'If Not IsNull(Me.txtLocator) Then
'WHEREStmt = WHEREStmt & "[Locator] = " & Me.txtLocator & " AND "
'If Not IsNull(Me.txtYear) Then
'WHEREStmt = WHEREStmt & "[Year] = " & Me.txtYear & " AND "

'Get rid of trailing AND statement
'WHEREStmt = Left(WHEREStmt, Len(WHEREStmt) - 5)

'DoCmd.OpenReport stDocName, acPreview, WHEREStmt
 
End Sub