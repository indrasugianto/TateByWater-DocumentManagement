' Component: Form_zClient Ledger OLD
' Type: document
' Lines: 439
' ============================================================

Option Compare Database
Private cls As New clsFormValidation


Private Sub cmdCaseList_Click()
    DoCmd.openform "frmYearWiseCaseList", acNormal, , , , , Me.cmbYear.value
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select Location"
        .show
        If .SelectedItems.Count <> 0 Then
            strPath = .SelectedItems(1)
            'Me.Text359 = strPath
            'Me.Check357 = -1
        Else
            ShowMessage "No location selected!"
        End If
    End With
ErrHandler:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmbYear_AfterUpdate()
    Me.FilterOn = False
    Me.Filter = "[Yr]= '" & Me.cmbYear & "'"
    Me.FilterOn = True
End Sub

Private Sub Command128_Click()
    lngChnc = 1
    DoCmd.RunCommand acCmdSaveRecord '=== TO SAVE ANY UNSAVED NEW CASE.
    DoCmd.openform "frmFamilyLaw", acNormal, , "CaseID=" & Me.CaseID, , acWindowNormal, Me.CaseID
End Sub

Private Sub Command353_Click()
    On Error GoTo ErrHandler_CmdClosingSheet
    DoCmd.OpenReport "rpt_Main_Closing", acViewPreview, , "[CaseID]=" & Me.CaseID
ErrHandler_CmdClosingSheet:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub Email_BeforeUpdate(Cancel As Integer)
    Dim ctlval As Control
    Set ctlval = Form.ActiveControl
    Cancel = EmailCheck(ctlval)
    Set ctlval = Nothing
End Sub



Private Sub OC_Email_BeforeUpdate(Cancel As Integer)
    Dim ctlval As Control
    Set ctlval = Form.ActiveControl
    Cancel = EmailCheck(ctlval)
    Set ctlval = Nothing
End Sub

'Private Sub Referral_AfterUpdate()
'  If Nz(Me.Referral.Column(1), "") = "Individual Referral" Then
'        Me.Individual_Referrer.Enabled = True
'        Me.Individual_Referrer.SetFocus
'    Else
'        Me.Individual_Referrer = Null
'        Me.Individual_Referrer.Enabled = False
'        Me.Referral.SetFocus
'    End If
'End Sub

Private Function EmailCheck(ByRef ctl As Control) As Boolean
    If ctl Like "*@*.*" Then
        ' its correct
        EmailCheck = False
    Else
        MsgBox "Please enter email address in correct format e.g., Test@human.com", vbOKOnly + vbInformation, "TB CMS: Incorrect Input"
        EmailCheck = True
    End If
End Function



'gaz
Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Sub FilterClear()
    'clear controls:
    Me.cmbYear = Null
    Me.chkClosed = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub chkClosed_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
        
    If Not IsNull(Me.chkClosed) Then strSQL = strSQL & " AND closed = " & Me.chkClosed
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

'Private Sub Text359_AfterUpdate()
'    If Nz(Me.Text359, "") = "" Then
'        Me.Check357 = 0
'    Else
'        Me.Check357 = -1
'    End If
'End Sub

'GH 2017-07-17:
'Private Sub CaseOpenDate_AfterUpdate()
'
'    'Debug.Print Form_frmCase.frmMatter.Form.Date2
'    'Debug.Print Form_frmCase.frmMatter.deta.Form.Detail.txtCaseOpenDate.value
'
'    On Error GoTo ErrHandler
'
'    Form_frmCase.Recordset.Edit
'    Form_frmCase.Recordset("CaseOpenDate") = Me.CaseOpenDate  '#11/11/1111#
'    Form_frmCase.Recordset.Update
'
'    Exit Sub
'ErrHandler:
'    If Err.Description = "No current record." Then
'        Me.Dirty = False
'        Form_frmCase.Recordset.AddNew
'        Form_frmCase.Recordset("CaseID") = Me.CaseID
'        Form_frmCase.Recordset("CaseOpenDate") = Me.CaseOpenDate  '#11/11/1111#
'        'Form_frmCase.Recordset.Update
'        Form_frmCase.Refresh
'    End If
'
'End Sub
'
'Private Sub Form_BeforeInsert(Cancel As Integer)
'    Call BeforeInsertRecord
'End Sub
'
'Sub BeforeInsertRecord()
'    If Me.Last_Name = "" Then
'        MsgBox "Last Name is mandatory."
'        Me.Last_Name.SetFocus
'        Me.Undo
'    Else
'        DoCmd.RunCommand acCmdSaveRecord
'        CaseGenerator
'    End If
'End Sub

Private Sub cmdSave_Click()
    'gaz 2017-07-23 disabled:

'    If Me.Dirty Then
'        cls.ExeCommand SaveRec 'gh: weird validation?
'    End If
'    CaseGenerator

    'gaz 2017-07-23
    
End Sub

Private Sub cmdAddNew_Click()
    'cls.ExeCommand Addrec
    DoCmd.GoToRecord , , acNewRec
End Sub


'----------------------------------
'GH Temporary Disabled functions
'Private Sub Form_Load()
'    If Not IsDisableEvents Then
'        'cls.RequiredControls = Array(Me.Last_Name, Me.First_Name, Me.Referral, Me.HmPhone)
'        cls.RequiredControls = Array(Me.Last_Name)
'        Dim dt As Long
'        Set cls.Form = Me
'        dt = Format(Date, "yy")
'        Me.cmbYear.value = dt
'        cmbYear_AfterUpdate
'    Else
'        IsDisableEvents = False
'    End If
'End Sub





'****************************************************
'-------------------- Code from FrmCase:
'****************************************************




'Option Compare Database

Private Sub cmdStatemenTrustAccount_Click()
    On Error GoTo ErrHanlder_cmdStatementrustAccount
        StatementofTrustAccount_Filter = True
        DoCmd.OpenReport "Statement of Trust Account", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
ErrHanlder_cmdStatementrustAccount:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
    
End Sub

Private Sub Command225_Click()
On Error Resume Next
    DoCmd.RunCommand acCmdRecordsGoToNew
End Sub

'gaz 2017-05-02
Private Sub cmdInvoice_Click()
    
    DoCmd.Close acReport, "New Invoice", acSaveYes
    
    intCaseID = Nz(Me.CaseID, 0)
    intMatterID = DLookup("MatterID", "qry_get_MatterID_from_zero_balance", "CaseID=" & intCaseID & " AND RetBal=0")
    'strFilter = "[CaseID]=" & Nz(Me.CaseID, 0) & " AND MatterID>=" & intMatterID
    
    If IsNull(intMatterID) Then
        strMatterID = ""
    Else
        strMatterID = " AND MatterID>" & intMatterID
    End If
    
    strFilter = "[CaseID]=" & Nz(Me.CaseID, 0) & strMatterID
    [Report_New Invoice].Filter = strFilter
    [Report_New Invoice].OrderBy = "MatterID"
    DoCmd.OpenReport "New Invoice", acViewPreview, , strFilter
End Sub

'gaz 2017-05-05
Private Sub cmdPastDueInvoice_Click()
    On Error GoTo ErrHandler_cmdPastDueInvoice
    
    DoCmd.Close acReport, "Invoice - Past Due", acSaveYes
    
    If Nz(Me.CaseID, 0) = 0 Then
        ShowMessage "Please select a client."
    Else
        DoCmd.OpenReport "Invoice - Past Due", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0) '& " AND [Past Due] = " & -1
    End If
ErrHandler_cmdPastDueInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
End Sub

'gaz 2017-05-05
Private Sub cmdBalanceInvoice_Click()
    On Error GoTo ErrHandler_cmdBalanceInvoice
                
    DoCmd.Close acReport, "Invoice - No Balance Due", acSaveYes
    'DoCmd.OpenReport "Invoice - No Balance Due", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0) & " AND [chkBalanceDue] = " & 0
    DoCmd.OpenReport "Invoice - No Balance Due", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
        
ErrHandler_cmdBalanceInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
    
End Sub

'gaz 2017-05-05
Private Sub CmdFullHistoryInvoice_Click()
    
    On Error GoTo ErrHandler_cmdInvoice
    
    DoCmd.Close acReport, "Invoice", acSaveYes
    DoCmd.OpenReport "Invoice", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
ErrHandler_cmdInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
End Sub

'? gaz: deleted during moving the controls to the upper part
'Private Sub Text272_BeforeUpdate(Cancel As Integer)
      'If Not IsNull(Me.Text272) Then
        'If CLng(Right(Year(Me.Text272), 2)) <> Me.Parent.Combo231.value Then
        '    MsgBox "The date is not of the year selected in the Ledger year column!", vbCritical + vbInformation, "Invalid Input"
        '    Cancel = True
        'End If
     'Else
           ' MsgBox "Case open date cannot be left blank!", vbInformation, "Invalid Input!"
           ' Cancel = True
     'End If
'End Sub

Private Sub cmdCreateHourlyBill_Click()
    'NEW TK BILL
    On Error GoTo ErrHandler_cmdCreateHourlyBill
    
    'Check if this CaseID is in the Time Keeping table:
    Dim rs As DAO.Recordset
    If DCount("CaseID", "TB Time Keeping", "CaseID=" & Nz(Me.CaseID, 0)) = 0 Then
        Set rs = CurrentDb.OpenRecordset("TB Time Keeping", dbOpenDynaset, dbSeeChanges)
        rs.AddNew
        rs.Fields("CaseID") = Nz(Me.CaseID, 0)
        rs.Update
        rs.Close
        Set rs = Nothing
    Else
        'do  nothing
    End If
    DoCmd.openform "Time Keeping", acNormal, , "[CaseID]=" & Nz(Me.CaseID, 0), , , "Hide Filter"
    Call [Form_Time Keeping].cmdAddNew_Click
    '[Form_Time Keeping].cmbClients = Me.CaseID
    'Call [Form_Time Keeping].cmbClients_AfterUpdate
    
ErrHandler_cmdCreateHourlyBill:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdHourlyBillList_Click()
    'SHOW TK BILLS FORM
    
    If IsNull(CaseID) Then
        MsgBox "Please select a client", , "TB CMS"
        Exit Sub
    End If
    DoCmd.openform "Time Keeping", acNormal, , "[CaseID]=" & Nz(Me.CaseID, 0), , , "Hide Filter"
    
    'GH 2017-08-11 it was not compiling so I temporary disabled it
    '[Form_Time Keeping].cmbClients = Me.CaseID
    
    If Not IsNull(CaseID) Then
       ' Call [Form_Time Keeping].cmbclients_AfterUpdate
    End If
End Sub

Private Sub cmdPrintLabel_Click()
    On Error GoTo ErrHandler_cmdPrintLabel_Click
    
    answer = MsgBox("Would you like to print this label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_address_label").IsLoaded Then
            DoCmd.Close acReport, "rpt_address_label", acSaveNo
        End If
        
        DoCmd.OpenReport "rpt_address_label", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    
ErrHandler_cmdPrintLabel_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdSubmitToFamilyLaw_Click()
'    Dim strSQL As String
'    strSQL = "Insert into [Family Law - Divorce]  ([CaseID]) values (" & lngCaseID & ")"
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (strSQL)
    lngChnc = 1
      DoCmd.RunCommand acCmdSaveRecord '=== TO SAVE ANY UNSAVED NEW CASE.
    DoCmd.openform "frmFamilyLaw", acNormal, , "CaseID=" & lngCaseID, , acWindowNormal, lngCaseID
End Sub



'000000000000000
'Private Sub Form_AfterUpdate()
'    lngCaseID = Nz(Me.CaseID, 0)
'    'CaseGenerator '=== this can take time in case of large scale data
'
'    If IsNull(Me.Number_) Then
'
'        maxval = Nz(DMax("Number_", "tblCase", "Yr='" & intval & "'"), 0)
'    End If
'
'    On Error GoTo errorhandler
'    Me.Parent.frmCaseList.Requery
'    Exit Sub
'errorhandler:
'    MsgBox "An error was encountered: " & Err.Description
'
'End Sub
'
Private Sub Form_Current()

    lngCaseID = Nz(Me.CaseID, 0)

    Debug.Print "CaseID: " & GetCaseID

'    If DLookup("Balance", "qryMatter", "CaseID=" & GetCaseID()) > 0 Then
'        Me.frmBilling.Form.Controls("chkBalanceDue") = True
'    Else
'        Me.frmBilling.Form.Controls("chkBalanceDue") = False
'    End If
End Sub
'



'Private Sub Case_Letter_AfterUpdate()
    'Me.yr = Format(Date, "yy") 'Right(Year(Date), 2)
    'Me.Year = Format(Date, "yy") 'Right(Year(Date), 2)
'End Sub


'Private Sub Last_Name_AfterUpdate()
    '?
'End Sub

Private Sub CaseOpenDate_AfterUpdate()
    nrMax = Nz(DMax("Number_", "tblCase", "Yr='" & Me.yr & "'"), 0)
    Me.Number_ = nrMax + 1
End Sub

Private Sub OrigAtty_AfterUpdate()
    Me.FileNo.Requery
End Sub

Private Sub cmdClose_Click()
    cls.ExeCommand Cancelrec
    DoCmd.Close acForm, Me.Name, acSave
End Sub