' Component: Form_Time Keeping
' Type: document
' Lines: 1084
' ============================================================

Option Compare Database
'Private cls As New clsFormValidation

Dim totalHours As Double

Private Sub Bill_Paid_BeforeUpdate(Cancel As Integer)
    If Not IsNull(Me.Bill_Paid) Then
            If Not IsNull(Me.Bill_Sent) Then
                If Me.Bill_Paid < Me.Bill_Sent Then
                    MsgBox "Bill payment date can't be less than the bill generation date.", vbInformation, "TB CMS: Invalid Input"
                    Cancel = True
                End If
            Else
                MsgBox "Bill sent/prepared date is required first. Kindly Enter it after pressing escape.", vbInformation, "TB CMS: Missing Date"
                Cancel = True
            End If
    Else
        Mbox "Bill Paid Date", 1
        Cancel = True
    End If
End Sub

Private Sub Bill_Sent_BeforeUpdate(Cancel As Integer)
    If Not IsNull(Me.Bill_Sent) Then
        Cancel = DateVarifier(Me.Bill_Sent)
    Else
        Mbox "Bill Sent", 1
        Cancel = True
    End If
End Sub

Private Sub Check285_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub Check315_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub Check318_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub chkadvance_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub chktkexceeds_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub chkTKless_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub cmbTAtty_AfterUpdate()
    Me!frmTimeTableDetail.Requery
End Sub

Private Sub cmdBillingOpenInvoiceFolder_Click()
If Not OpenDocumentFolder(CaseID, "Client Invoices") Then
        MsgBox "Failed to open document folder...", , "TB CMS"
    End If
End Sub

Private Sub cmdCompFullHistory_Click()
  Dim sAttachmentName As String
    sAttachmentName = IANumber & ": " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
 
 On Error GoTo ErrHandler_CaseID_Click
    
    If Me.Bill_Closed = 0 Then
        MsgBox "Stmts / Invoices generated after TK closed at Take Off." & vbCrLf & "To Preview click TK Statement." & vbCrLf & "Please see Admin (KV/PM) for early TK Close.", vbInformation, "TB CMS"
    Else
    
    If Me.InvoiceTotalAdvance = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceADV"
    End If
    
    If Me.StatementLessTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceStmt"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx1"
    End If
    
    If Me.InvoiceAdvCostFee = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx2"
    End If
    
    If Me.InvoiceCostHold = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx3Costs"
    End If
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
    Reports(sExistingReportName).Caption = sAttachmentName
End If

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
   
End Sub

Private Sub cmdCompShort_Click()

 Dim sAttachmentName As String
    sAttachmentName = IANumber & ": " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
 
 On Error GoTo ErrHandler_CaseID_Click
    
    If Me.Bill_Closed = 0 Then
        MsgBox "Stmts / Invoices generated after TK closed at Take Off." & vbCrLf & "To Preview click TK Statement." & vbCrLf & "Please see Admin (KV/PM) for early TK Close.", vbInformation, "TB CMS"
    Else
    
    If Me.InvoiceTotalAdvance = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceADVS"
    End If
    
    If Me.StatementLessTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceStmtS"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx1S"
    End If
    
    If Me.InvoiceAdvCostFee = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx2S"
    End If
    
    If Me.InvoiceCostHold = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx3CostsS"
    End If
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
    Reports(sExistingReportName).Caption = sAttachmentName
End If

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
   
End Sub

Private Sub cmdEmailLong_Click()
  Dim sAttachmentName As String
    sAttachmentName = IANumber & ": " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
 
 On Error GoTo ErrHandler_CaseID_Click
    
    If Me.Bill_Closed = 0 Then
        MsgBox "Stmts / Invoices generated after TK closed at Take Off." & vbCrLf & "To Preview click TK Statement." & vbCrLf & "Please see Admin (KV/PM) for early TK Close.", vbInformation, "TB CMS"
    Else
    
    If Me.InvoiceTotalAdvance = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceADV"
    End If
    
    If Me.StatementLessTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceStmt"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx1"
    End If
    
    If Me.InvoiceAdvCostFee = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx2"
    End If
    
    If Me.InvoiceCostHold = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx3Costs"
    End If
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If

DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
    Reports(sExistingReportName).Caption = sAttachmentName
DoCmd.SendObject acSendReport, , acFormatPDF, Nz(Me.Email), , , "Invoice / Statement " & sAttachmentName, Nz(Title) & " " & Nz(Last_Name) & "," & vbLf & vbLf & "Please see the attached invoice. To make payment, please visit www.tatebywater.com/make-a-payment, send a check payable to Tate Bywater, or call the office at (703)938-5100. We appreciate your prompt payment." & vbLf & vbLf & "Thank you.", True, "strReportName" & Date
   

End If

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description

End Sub

Private Sub cmdEmailShort_Click()
 Dim sAttachmentName As String
    sAttachmentName = IANumber & ": " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
 
 On Error GoTo ErrHandler_CaseID_Click
    
    If Me.Bill_Closed = 0 Then
        MsgBox "Stmts / Invoices generated after TK closed at Take Off." & vbCrLf & "To Preview click TK Statement." & vbCrLf & "Please see Admin (KV/PM) for early TK Close.", vbInformation, "TB CMS"
    Else
    
    If Me.InvoiceTotalAdvance = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceADVS"
    End If
    
    If Me.StatementLessTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceStmtS"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx1S"
    End If
    
    If Me.InvoiceAdvCostFee = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx2S"
    End If
    
    If Me.InvoiceCostHold = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx3CostsS"
    End If
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
    
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
    Reports(sExistingReportName).Caption = sAttachmentName
DoCmd.SendObject acSendReport, , acFormatPDF, Nz(Me.Email), , , "INV/STMT " & sAttachmentName, Nz(Title) & " " & Nz(Last_Name) & "," & vbLf & vbLf & "Please see the attached invoice. To make payment, please visit www.tatebywater.com/make-a-payment, send a check payable to Tate Bywater, or call the office at (703)938-5100. We appreciate your prompt payment." & vbLf & vbLf & "Thank you.", True, "strReportName" & Date
   

End If

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
    
End Sub

Private Sub cmdPrevStatement_Click()
'    On Error GoTo ErrHandler_CaseID_Click
'
'    If Me.Bill_Closed = 0 Then
'        MsgBox "Statements and Invoices cannot be generated until TK is closed." & vbCrLf & "TKs are closed at Take Off.", vbInformation, "TB CMS"
'    Else
'
'    If Me.InvoiceTotalAdvance = -1 Then
'        strReportName = "rpt_TKTotalAdvance"
'    End If
'
'    If Me.InvoiceExceedsTrust = -1 Then
'        strReportName = "rpt_TKExceedsTrust"
'    End If
'
'    If Me.StatementLessTrust = -1 Then
'        strReportName = "rpt_TKLessTrust"
'    End If
'
'     If CurrentProject.AllReports(strReportName).IsLoaded Then
'        DoCmd.Close acReport, strReportName, acSaveNo
'    End If
'    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
'End If
'
'ErrHandler_CaseID_Click:
'    If Err.Number <> 0 Then ShowMessage Err.Description
           
End Sub
Private Sub cmdPrevHourlyInvoice_Click()

End Sub

Private Sub cmdRecordShortTK_Click()
         
    Dim strReportName As String
    Dim CaseID As Variant
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim strSQL As String
    
    CaseID = Forms!frmClientLedger!CaseID
         
    If pcaempty(CaseID) Then
        MsgBox "Please select client."
        Exit Sub
    End If
         
    If Me.Bill_Closed = 0 Then
        MsgBox "Statements and Invoices cannot be generated until TK is closed." & vbCrLf & "Record TK after TK closed.", vbInformation, "TB CMS"
    Else
    
        strSQL = "Insert into [tbl_InvoiceSent] "
                strSQL = strSQL & "        (CaseID, InvBalance, TKDate, [TK Sent], TKNumber)"
    '            strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
                strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", " & pcaConvertEmpty(Me.txtAR, 0) & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
        Form_frmInvoiceSent.Requery
        
    
        strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
    
        'Me.Requery
        Me.Bill_Sent.Requery
    
  
        DocumentFileName = IANumber & "- " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
        DocumentFileName = Replace(DocumentFileName, "/", "-")
         AllInvoicesFolderName = GetAllInvoicesFolderName(Form_frmClientLedger.CaseID)
        DocumentFolderName = GetDocumentFolderName(CaseID, "Client Invoices")
    
 
        If Me.Bill_Closed = 0 Then
            MsgBox "Stmts / Invoices generated after TK closed at Take Off." & vbCrLf & "To Preview click TK Statement." & vbCrLf & "Please see Admin (KV/PM) for early TK Close.", vbInformation, "TB CMS"
        Else
        
        If Me.InvoiceTotalAdvance = -1 Then
          strReportName = "rpt_Comprehensive_InvoiceADVS"
        End If
        
        If Me.StatementLessTrust = -1 Then
          strReportName = "rpt_Comprehensive_InvoiceStmtS"
        End If
        
        If Me.InvoiceExceedsTrust = -1 Then
          strReportName = "rpt_Comprehensive_InvoiceTKEx1S"
        End If
        
        If Me.InvoiceAdvCostFee = -1 Then
          strReportName = "rpt_Comprehensive_InvoiceTKEx2S"
        End If
        
        If Me.InvoiceCostHold = -1 Then
          strReportName = "rpt_Comprehensive_InvoiceTKEx3CostsS"
        End If

        
        If CurrentProject.AllReports(strReportName).IsLoaded Then
            DoCmd.Close acReport, strReportName, acSaveNo
        End If
         
        DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID, acHidden
        If FolderExistsCreate(DocumentFolderName, True) Then
            DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, DocumentFolderName & DocumentFileName & ".pdf"
            
            
            'save the document to _ALL INVOICES folder
            If FolderExistsCreate(AllInvoicesFolderName, True) Then
                DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, AllInvoicesFolderName & DocumentFileName & ".pdf"
            End If
            
            DoCmd.Close acReport, strReportName, acSaveNo
            
            If Not SaveCaseDocument(CaseID, "Client Invoices", DocumentFolderName & DocumentFileName & ".pdf") Then
                MsgBox "Fail to save Case Document record..."
            End If
            
        End If
    
    End If
    
    
    'Call cmdPreview_Click
    'On Error GoTo ErrHandler_CaseID_Click
    
    'If Me.Discount > 0 Then
       ' strReportName = "Invoice Attach - Hourly w Discount"
    'Else
'        strReportName = "Invoice Attach - Hourly"
'    'End If
'
'    If CurrentProject.AllReports(strReportName).IsLoaded Then
'        DoCmd.Close acReport, strReportName, acSaveNo
'    End If
'
'    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
  
ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
    'Me.[Bill Sent].Requery
    'Me.Requery
    MsgBox "Recorded", , "TB CMS"
     End If

End Sub

Private Sub cmdTest_Click()
    If Me.Bill_Closed = 0 Then
        DoCmd.openform "frmAdminLoginTK", acNormal
    Else
        MsgBox "TK is already closed.", vbInformation, "TB CMS"
    End If
End Sub

Private Sub cmdTotalTKs_Click()
     On Error GoTo ErrHandler_CaseID_Click

        strReportName = "rptComprehensiveTKStatement"
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If

    DoCmd.OpenReport strReportName, acViewPreview, , "[CaseID]=" & Me.CaseID

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub


Private Sub Form_AfterUpdate()
    lngBill_ID = Nz(Me.Bill_ID, 0)
End Sub

'gaz 2017-05-05
'Sub setIANumber_OLD()
'    strCaseID = [Form_Client Ledger].CaseID
'
'    If Not IsNull(strCaseID) Then
'        currNr = DLookup("CountOfIANumber", "qry_get_time_keeping_numbers", "CaseID = " & strCaseID)
'        strIANumber = "TK-" & Nz(currNr) + 1
'    End If
'
'    If IsNull(IANumber) Then IANumber = strIANumber
'
'    'also set the Discount to 0
'    If IsNull(Discount) Then Discount = 0
'
'End Sub

'gaz 2017-05-07

'Private Sub Command225_Click()
'    If Me.Dirty Then
'        cls.ExeCommand SaveRec
'    End If
'    Me.cmbBills.Requery
'End Sub

'Private Sub Command227_Click()
'    cls.ExeCommand Cancelrec
'    DoCmd.Close acForm, Me.Name, acSaveNo
'End Sub

Sub resetTimerControls()
    Me.TimerInterval = 0
    Me.txtTimeStart = 0
    Me.txtTimeStop = 0
    Me.txtTimeTotal = 0
    Me.txtTimeTotalHours = 0
End Sub

Private Sub cmdTimeStart_Click()
    Me.TimerInterval = 1000
    Me.txtTimeStart = Format(Now(), "hh:mm:ss")
End Sub

Private Sub cmdTimeStop_Click()
    Me.TimerInterval = 0
    Call UpdateTime
End Sub

Private Sub Form_Timer()
    Debug.Print Now()
    Call UpdateTime
End Sub

Sub UpdateTime()
    Me.txtTimeStop = Format(Now(), "hh:mm:ss")
    Me.txtTimeTotal = DateDiff("n", Me.txtTimeStart, Me.txtTimeStop)
    txtTimeTotalHours = txtTimeTotal / 60
    totalHours = txtTimeTotal / 60
End Sub


Private Sub cmdPrintInvoice_Click()
    On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
    If Me.Discount > 0 Then
        strReportName = "Invoice Attach - Hourly w Discount"
    Else
        strReportName = "Invoice Attach - Hourly"
    End If
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports(strReportName).IsLoaded Then
            DoCmd.Close acReport, strReportName, acSaveNo
        End If
        
        DoCmd.OpenReport strReportName, acNormal, , "[Bill_ID]=" & Me.Bill_ID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description

End Sub

'Sub cmbclients_AfterUpdate()
'    'Me.cmbBills.RowSourceType = "Table/Query"
'
'    'GH 2017-08-11 it was not compiling so I temporary disabled it
'    'Me.cmbBills.RowSource = "select * from qryBillList where CaseID=" & Me.cmbClients
'
'    Me.cmbBills = Me.cmbBills.Column(0, 0)
'    Call cmbBills_AfterUpdate
'End Sub

Private Sub Form_Load()
'    Set cls.Form = Me
'    If Nz(Me.OpenArgs, "") <> "" Then
'        If Me.OpenArgs = "Hide Filter" Then
'            Me.cmbbills.Visible = False
'            Me.NavigationButtons = False
'        End If
'    Else
'        Me.cmbbills.Visible = True
'        Me.NavigationButtons = True
'    End If

    'gaz
    Call resetTimerControls
    'Call setIANumber
    
End Sub

Private Sub cmdRecordTKStatement_Click()
         
    Dim strReportName As String
    Dim CaseID As Variant
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim strSQL As String
    
    CaseID = Forms!frmClientLedger!CaseID
         
    If pcaempty(CaseID) Then
        MsgBox "Please select client.", , "TB CMS"
        Exit Sub
    End If
         
    If Me.Bill_Closed = 0 Then
        MsgBox "Statements and Invoices cannot be generated until TK is closed." & vbCrLf & "Record TK after TK closed.", vbInformation, "TB CMS"
    Else
    
        strSQL = "Insert into [tbl_InvoiceSent] "
                strSQL = strSQL & "        (CaseID, InvBalance, TKDate, [TK Sent], TKNumber)"
    '            strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
                strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", " & pcaConvertEmpty(Me.txtAR, 0) & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
        Form_frmInvoiceSent.Requery
        
    
        strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
    
        'Me.Requery
        Me.Bill_Sent.Requery
    
  
        DocumentFileName = IANumber & "- " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
        DocumentFileName = Replace(DocumentFileName, "/", "-")
        AllInvoicesFolderName = GetAllInvoicesFolderName(Form_frmClientLedger.CaseID)
        DocumentFolderName = GetDocumentFolderName(CaseID, "Client Invoices")
    
 
        If Me.Bill_Closed = 0 Then
            MsgBox "Stmts / Invoices generated after TK closed at Take Off." & vbCrLf & "To Preview click TK Statement." & vbCrLf & "Please see Admin (KV/PM) for early TK Close.", vbInformation, "TB CMS"
        Else
        
        If Me.InvoiceTotalAdvance = -1 Then
            strReportName = "rpt_Comprehensive_InvoiceADV"
        End If
        
        If Me.StatementLessTrust = -1 Then
            strReportName = "rpt_Comprehensive_InvoiceStmt"
        End If
        
        If Me.InvoiceExceedsTrust = -1 Then
            strReportName = "rpt_Comprehensive_InvoiceTKEx1"
        End If
        
        If Me.InvoiceAdvCostFee = -1 Then
            strReportName = "rpt_Comprehensive_InvoiceTKEx2"
        End If
        
        If Me.InvoiceCostHold = -1 Then
            strReportName = "rpt_Comprehensive_InvoiceTKEx3Costs"
        End If
        
        If CurrentProject.AllReports(strReportName).IsLoaded Then
            DoCmd.Close acReport, strReportName, acSaveNo
        End If
         
        DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID, acHidden
        If FolderExistsCreate(DocumentFolderName, True) Then
            DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, DocumentFolderName & DocumentFileName & ".pdf"
            
            
             'save the document to _ALL INVOICES folder
            If FolderExistsCreate(AllInvoicesFolderName, True) Then
                DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, AllInvoicesFolderName & DocumentFileName & ".pdf"
            End If
            
            DoCmd.Close acReport, strReportName, acSaveNo
            
            If Not SaveCaseDocument(CaseID, "Client Invoices", DocumentFolderName & DocumentFileName & ".pdf") Then
                MsgBox "Failed to save Case Document record...", , "TB CMS"
            End If
            
        End If
    
    End If
    
    
    'Call cmdPreview_Click
    'On Error GoTo ErrHandler_CaseID_Click
    
    'If Me.Discount > 0 Then
       ' strReportName = "Invoice Attach - Hourly w Discount"
    'Else
'        strReportName = "Invoice Attach - Hourly"
'    'End If
'
'    If CurrentProject.AllReports(strReportName).IsLoaded Then
'        DoCmd.Close acReport, strReportName, acSaveNo
'    End If
'
'    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
  
ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
    'Me.[Bill Sent].Requery
    'Me.Requery
    MsgBox "Recorded", , "TB CMS"
     End If
End Sub

Private Sub cmdCreateAR_Click()
    
'    strSQL = "Insert into [Matter and AR] "
'            strSQL = strSQL & "        (CaseID, Date2, Pay_Outlay, Charge)"
'            strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, 'Advanced Legal Work (" & Nz(Me.IANumber, 0) & ")', " & Me.txtTotalBill & ")"
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL
'
'
'    strSQL = "Update [TB Time Keeping] set TKLocked = -1 where IANumber = '" & Me.IANumber & "'"
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL
'
'    'lock controls:
'    Me.cmdCreateAR.Enabled = False
'    Form_frmTimeTableDetail.AllowAdditions = False
'    Me.cmdInsertTime.Enabled = False
'
'    Form_frmMatter.Requery
'    MsgBox "AR Inserted", , "TB CMS"
'    strSQL = "Insert into [Matter and AR] (CaseID, Charge, Date2, Pay_Outlay)"
'        strSQL = strSQL & " values (" & _
'                          Me.txtCaseID & ", " & _
'                          Nz(Me.txtEarned, 0) & ", #" & _
'                          Form_frmTakeOff.txtTakeOffDate & "#, 'Advanced Legal Work (TK)')"
    
    
End Sub


'Private Sub cmbtenths_Click()
'    Me![txttenths].Visible = yes   'make Text Box invisible
'    Me.TimerInterval = 5000         'set Timer Interval to 5 seconds
'    Me![txttenths].Visible = No
'End Sub



Private Sub cmdPreview_Click()

    On Error GoTo ErrHandler_CaseID_Click

        strReportName = "Invoice Attach - Hourly"
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If

    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
       
End Sub

Private Sub cmdInsertTime_Click()

'    strSQL = "Insert into [TB Time Keeping] (CaseID, IANumber)"
'    strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", '" & strIANumber & "')"
'
'    Debug.Print "strSQL=" & strSQL
'    CurrentDb.Execute strSQL

    Debug.Print totalHours
    totalHours = Round(totalHours, 2)
    
    strSQL = "Insert into [tblTimeTableDetail] "
            strSQL = strSQL & "        (Bill_ID, Time_)"
            strSQL = strSQL & " values (" & Form_frmTimeTableDetail.Bill_ID & ", " & totalHours & ")" 'Me.txtTimeTotalHours
    Debug.Print strSQL
    CurrentDb.Execute strSQL
    
    MsgBox "Time Inserted", , "TB CMS"
    Me.Requery
End Sub


'gaz 2017-07-31
Sub cmdAddNew_Click() 'add record
    If Me.Dirty Then Me.Dirty = False
    'check if this is the first TK.
    strCaseID = Form_frmClientLedger.CaseID
    TKcount = DCount("Bill_ID", "[TB Time Keeping]", "CaseID=" & strCaseID)
    If TKcount > 0 Then
        'if this is not the first TK then check if all previous TKs are closed
        TKcountClosed = DCount("Bill_ID", "[TB Time Keeping]", "CaseID=" & strCaseID & " and [Bill Closed]=True")
        If TKcount = TKcountClosed Then
            'all TKs are closed, Good to go.
            'Debug.Print "OK"
            Call addNewTK(strCaseID)
        Else
            'error - can't go further.
            'Debug.Print "Error"
            MsgBox "You cannot add a new TK until the current TK is closed.", vbInformation, "TB CMS"
        End If
    Else
        'making is the first TK:
        Call addNewTK(strCaseID)
    End If
End Sub

Sub addNewTK(strCaseID)
    If Not IsNull(strCaseID) Then
        currNr = DLookup("CountOfIANumber", "qry_get_time_keeping_numbers", "CaseID = " & strCaseID)
        strIANumber = "TK-" & Nz(currNr) + 1
    End If
            
    strSQL = "Insert into [TB Time Keeping] (CaseID, [Bill Open], Discount, IANumber)"
    strSQL = strSQL & " values (" & strCaseID & ", #" & Date & "#, 0, '" & strIANumber & "')"
    
    Debug.Print "strSQL=" & strSQL
    CurrentDb.Execute strSQL, dbSeeChanges
    
    'Debug.Print "strIANumber=" & strIANumber
    
    Me.cmbBills.RowSource = "select * from qryBillList where CaseID=" & strCaseID
    Me.cmbBills.Requery
    
    maxBillID = DMax("Bill_ID", "qryBillList", "CaseID=" & strCaseID)
    Me.cmbBills = maxBillID
    
    Call cmbBills_AfterUpdate
    
    MsgBox "New TK added.", , "TB CMS"
End Sub


    'Me.Requery
    'Me.Recordset.MoveLast
    'Me.filter = "[CaseID]=" & strCaseID & " and IANumber = '" & strIANumber & "'"
    'Me.FilterOn = True
    
'On Error GoTo errorhandler
    'If IsNull(IANumber) Then IANumber = strIANumber
    
    'also set the Discount to 0
    'If IsNull(Discount) Then Discount = 0

    '[Form_Time Keeping].Requery
    '[Form_Time Keeping].Recordset.MoveLast
    'Me.cmbBills.Requery

    'strCaseID = CaseID
    'DoCmd.GoToRecord , , acNewRec
    'CaseID = strCaseID
    'Call setIANumber
    'Me.Bill_Sent = Date
            
    'Me.filter = "[CaseID]=" & strCaseID
    'Me.FilterOn = True
    
'    Exit Sub
'errorhandler:
'    If Err.Description <> "Cannot add record(s); join key of table 'TB Time Keeping' not in recordset." Then
'        MsgBox Err.Description
'    End If

Private Sub cmdCloseBill_Click()
 '   MsgBox "Ok"
'    MsgBox "TK is already closed.", vbCritical, "TB CMS"
'    If Me.Bill_Closed = 0 Then
'        DoCmd.openform "frmAdminLoginTK", acNormal
'    Else
'        MsgBox "TK is already closed.", vbCritical, "TB CMS"
'    End If
'    answer = MsgBox("Are you sure you want to close this TK?" & vbCrLf & "This action cannot be undone.", vbYesNo, "TB CMS")
'    If answer = vbYes Then
'        Me.[BilL Closed Date] = Date
'        Me.Bill_Closed = True
'        disableCloseTK
'    End If
End Sub

Sub disableCloseTK()
    Me.Bill_Closed.Enabled = False
    Me.cmdCloseBill.Enabled = False
    Me.txtBillClosedDate.Enabled = False
End Sub

Sub enableCloseTK()
    Me.Bill_Closed.Enabled = True
    Me.cmdCloseBill.Enabled = True
    Me.txtBillClosedDate.Enabled = True
End Sub

Private Sub Form_Current()
    lngBill_ID = Nz(Me.Bill_ID, 0)
        
    Me.cmbBills.RowSource = "select * from qryBillList where CaseID=" & Form_frmClientLedger.CaseID
    Me.cmbBills.Requery
    
    'lock controls:
    Me.cmdCreateAR.Enabled = Not Nz(TKLocked, "0")
    Form_frmTimeTableDetail.AllowAdditions = Not Nz(TKLocked, "0")
    Form_frmTimeTableDetail.AllowEdits = Not Nz(TKLocked, "0")
    Form_frmTimeTableDetail.AllowDeletions = Not Nz(TKLocked, "0")
    Me.cmdInsertTime.Enabled = Not Nz(TKLocked, "0")
End Sub

Sub cmbBills_AfterUpdate()
    If Not IsNull(Me.cmbBills) Then
        Me.Filter = "Bill_ID=" & Me.cmbBills
    Else
        Me.Filter = "Bill_ID=0"
    End If
    Me.FilterOn = True
        
    If Me.Bill_Closed = True Then
        Call disableCloseTK
    Else
        Call enableCloseTK
    End If
End Sub



Private Sub cmdCompCurrent_Click()
    
    'Debug.Print Me.CaseID
    On Error GoTo ErrHandler_CaseID_Click
    
    'DoCmd.OpenReport "rptInvoiceComprARCur", acViewPreview, , "[CaseID]=" & Me.CaseID
    'DoCmd.OpenReport "rptInvoiceComprARCur", acViewPreview, , "[CaseID]=" & Me.CaseID
    
    'parent reports:
    'rpt_Compr_InvoiceADVCur
    'rpt_Compr_InvoiceStmtCur
    'rpt_Compr_InvoiceTKExCur
    
    'subreports / queries
    'rptInvoiceComprARCur / qry_InvoiceAR_curr
    'rptInvoiceComprPymtsARCur / qry_InvoicePymts_curr
    'rptInvoiceComprTrustCur / qryInvoiceComprehensiveTrustCredit
    
    'Exit Sub
'
'     If fncIsThereAZeroInMatterARForm = False Then
'        MsgBox "Please use the Comprehensive Full Invoice.", , "TB CMS"
'        Exit Sub
'    End If
    
    If Me.Bill_Closed = 0 Then
        MsgBox "Statements and Invoices cannot be generated until TK is closed." & vbCrLf & "TKs are closed at Take Off.", vbInformation, "TB CMS"
    Else
        If Me.InvoiceTotalAdvance = -1 Then
            strReportName = "rpt_Compr_InvoiceADVCur"
        End If
        
        If Me.InvoiceExceedsTrust = -1 Then
            strReportName = "rpt_Compr_InvoiceTKExCur"
        End If
        
        If Me.StatementLessTrust = -1 Then
            strReportName = "rpt_Compr_InvoiceStmtCur"
        End If
        
        If CurrentProject.AllReports(strReportName).IsLoaded Then
            DoCmd.Close acReport, strReportName, acSaveNo
        End If
        
        DoCmd.OpenReport strReportName, acViewPreview, , "Bill_ID=" & Me.Bill_ID
        
    End If

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub Text287_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub Text289_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub Text291_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub Text293_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub Text295_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub

Private Sub TrustatClose_Click()
    If Me.Bill_Closed = 0 Then
        MsgBox "Admin Access Only.  TK must be closed before updating.", vbInformation, "TB CMS"
        Me.Undo
    Else
    
    If Me.Bill_Closed = -1 Then
        answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
        If answer = "tb2740" Then
    Else
    Me.Undo
    End If
    End If
    End If
End Sub