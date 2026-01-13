' Component: Form_frm_invoices_summary
' Type: document
' Lines: 487
' ============================================================

Option Compare Database

Private Sub CaseID_Click()
    On Error GoTo ErrHandler_CaseID_Click

    If CurrentProject.AllReports("Invoice").IsLoaded Then
        DoCmd.Close acReport, "Invoice", acSaveNo
    End If

    DoCmd.OpenReport "Invoice", acViewPreview, , "[CaseID]=" & Me.CaseID

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
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

Private Sub chklongterm_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClearFilter_Click()
    Call FilterClear
End Sub

Private Sub cmbclose_Click()
    DoCmd.Close acForm, Me.Name
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

Private Sub Cmd_PreviewNew_Click()
   
'   If fncIsThereAZeroInMatterARForm = False Then
'        MsgBox "Please use the Full History Invoice."
'        Exit Sub
'    End If

   DoCmd.Close acReport, "New Invoice", acSaveYes
    
    intCaseID = Nz(Me.CaseID, 0)
    'intMatterID = DLookup("MatterID", "qry_get_MatterID_from_zero_balance", "CaseID=" & intCaseID & " AND RetBal=0")
    'strFilter = "CaseID=" & intCaseID & " AND Balance=0 order by OrderNr desc"
    strFilter = "CaseID=" & intCaseID & " AND Balance=0"
    intOrderNr = DMax("OrderNr", "qry_current_invoice", strFilter)
    'strFilter = "[CaseID]=" & Nz(Me.CaseID, 0) & " AND MatterID>=" & intMatterID
    
    If IsNull(intOrderNr) Then
        strOrderNr = ""
    Else
        strOrderNr = " AND OrderNr>" & intOrderNr
    End If
    
    strFilter = "[CaseID]=" & Nz(Me.CaseID, 0) & strOrderNr
    [Report_New Invoice].Filter = strFilter
    
    [Report_New Invoice].OrderBy = "OrderNr"
    
    MsgBox "Before printing make sure there is a fee at the top of the invoice.", , "TB CMS"
    DoCmd.OpenReport "New Invoice", acViewPreview, , strFilter
End Sub

Private Sub Cmd_PrintNew_Click()
   On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("New Invoice").IsLoaded Then
            DoCmd.Close acReport, "New Invoice", acSaveNo
        End If
        
        DoCmd.OpenReport "New Invoice", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
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

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdLTC_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdPast90_Click()
     dt30minus = DateAdd("d", "-90", Date)
    
    Dim strSQL As String
    strSQL = "[Balance Due Date] <= #" & dt30minus & "#"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub cmdPastDueLT_Click()
If [Balance Due Date] > Date Then
    MsgBox "This Invoice is Not Past Due.", , "TB CMS"
Else
On Error GoTo ErrHandler_cmdPastDue_Click
    
    If CurrentProject.AllReports("Invoice - Past Due").IsLoaded Then
        DoCmd.Close acReport, "Invoice - Past Due", acSaveYes
    End If
    
    DoCmd.OpenReport "Invoice - Past Due", acViewPreview, , "[CaseID]=" & Me.CaseID
    
ErrHandler_cmdPastDue_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End If
End Sub

Private Sub cmdPrintInvoice_Click()
    On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("Invoice").IsLoaded Then
            DoCmd.Close acReport, "Invoice", acSaveNo
        End If
        
        DoCmd.OpenReport "Invoice", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub
Private Sub cmdPastDue_Click()
If [Balance Due Date] > Date Then
    MsgBox "This Invoice is Not Past Due.", , "TB CMS"
Else
On Error GoTo ErrHandler_cmdPastDue_Click
    
    If CurrentProject.AllReports("Invoice - Past Due").IsLoaded Then
        DoCmd.Close acReport, "Invoice - Past Due", acSaveYes
    End If
    
    DoCmd.OpenReport "Invoice - Past Due", acViewPreview, , "[CaseID]=" & Me.CaseID
    
ErrHandler_cmdPastDue_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End If
End Sub
Private Sub cmdPreview_Click()
    On Error GoTo ErrHandler_CaseID_Click
    
    If CurrentProject.AllReports("Invoice").IsLoaded Then
        DoCmd.Close acReport, "Invoice2", acSaveNo
    End If
    
    DoCmd.OpenReport "Invoice2", acViewPreview, , "[CaseID]=" & Me.CaseID
    
ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub
Sub FilterClear()
    'clear controls:
    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.txtClient = Null
    Me.chkNonZero = Null
    Me.cmbAssoc = Null
    Me.cmbCodeVal = Null
    Me.cmdLTC = Null
    
    Me.Filter = ""
    Me.FilterOn = False
    DoCmd.ApplyFilter , "BalRetCalculated <>0"
End Sub
Private Sub cmdPastDuePrint_Click()
    On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("Invoice - Past Due").IsLoaded Then
            DoCmd.Close acReport, "Invoice - Past Due", acSaveNo
        End If
    
        DoCmd.OpenReport "Invoice - Past Due", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdNoBalancePrint_Click()
    On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("Invoice - No Balance Due").IsLoaded Then
            DoCmd.Close acReport, "Invoice - No Balance Due", acSaveNo
        End If
        
        DoCmd.OpenReport "Invoice - No Balance Due", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub chkNonZero_AfterUpdate()
    Call FilterMe
End Sub
Private Sub cmdPrintLabel_Click()
 On Error GoTo ErrHandler_cmdPrintLabelP_Click
    
    If IsNull(Me.Executor) Then
    
    answer = MsgBox("Would you like to print this label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_address_label").IsLoaded Then
            DoCmd.Close acReport, "rpt_address_label", acSaveNo
        End If
        
        DoCmd.OpenReport "rpt_address_label", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    End If
   
    If Not IsNull(Me.Executor) Then
    
    answer = MsgBox("Would you like to print this label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_address_labelEx").IsLoaded Then
            DoCmd.Close acReport, "rpt_address_labelEx", acSaveNo
        End If
        
        DoCmd.OpenReport "rpt_address_labelEx", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    End If
ErrHandler_cmdPrintLabelP_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub
Private Sub cmdPrintPastDueLT_Click()
 On Error GoTo ErrHandler_cmdPrintInvoice_Click
    
    answer = MsgBox("Would you like to print this report?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("Invoice - Past Due").IsLoaded Then
            DoCmd.Close acReport, "Invoice - Past Due", acSaveNo
        End If
    
        DoCmd.OpenReport "Invoice - Past Due", acNormal, , "[CaseID]=" & Me.CaseID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub
Private Sub cmdRecordInvoice_Click()
    Dim sExistingReportName As String
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim strSQL As String
    Dim sFileFullPath As String

    
    If pcaempty(CaseID) Then
        MsgBox "Please select a client", , "TB CMS"
    Else
       

        
        
        'put the Invoice PDF copy to \Invoices folder

        sExistingReportName = "Invoice2"
        DocumentFileName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Invoice" & " " & Date
        DocumentFileName = Replace(DocumentFileName, "/", "-")
        
        If Closed Then
            DocumentFolderName = GetClosedDocumentFolderName(CaseID, "Client Invoices")
        Else
            DocumentFolderName = GetDocumentFolderName(CaseID, "Client Invoices")
        End If
        
        AllInvoicesFolderName = GetAllInvoicesFolderName(CaseID)
        

        DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(CaseID, 0), acHidden  '& " AND [Past Due] = " & -1
        If FolderExistsCreate(DocumentFolderName, True) Then
            'save the document to client document invoice folder
            DoCmd.OutputTo acOutputReport, sExistingReportName, acFormatPDF, DocumentFolderName & DocumentFileName & ".pdf"
            
            'save the document to _ALL INVOICES folder
            If FolderExistsCreate(AllInvoicesFolderName, True) Then
                DoCmd.OutputTo acOutputReport, sExistingReportName, acFormatPDF, AllInvoicesFolderName & DocumentFileName & ".pdf"
            End If
            
            DoCmd.Close acReport, sExistingReportName, acSaveNo
            
            If Not SaveCaseDocument(CaseID, "Client Invoices", DocumentFolderName & DocumentFileName & ".pdf") Then
                MsgBox "Failed to save Case Document record...", , "TB CMS"
            End If
                        
            strSQL = "Insert into [tbl_InvoiceSent] "
                    strSQL = strSQL & "        (CaseID, InvSent, InvBalance)"
                    strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, " & Me.BalRetCalculated & ")"
            Debug.Print strSQL
            CurrentDb.Execute strSQL, dbSeeChanges
            Me.Requery
            
            MsgBox "Invoice Sent Recorded", , "TB CMS"

        End If
    End If
End Sub
Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '*" & Me.txtClient & "*'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.cmbCodeVal) Then strSQL = strSQL & " AND CodeVal = '" & Me.cmbCodeVal & "'"
    If Not IsNull(Me.cmdLTC) Then strSQL = strSQL & " AND [Long Term Collections] = " & IIf(Me.cmdLTC, 1, 0)
    If Not IsNull(Me.chkNonZero) Then
        If chkNonZero Then
            strSQL = strSQL & " AND BalRetCalculated <>0 "
        Else
            strSQL = strSQL & " AND BalRetCalculated = 0 "
        End If
    End If

    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub
Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub
Private Sub cmdPast30_Click()
    
    dt30minus = DateAdd("d", "-30", Date)
    
    Dim strSQL As String
    strSQL = "[Balance Due Date] <= #" & dt30minus & "#"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub
Private Sub cmdPast180_Click()
    dt180minus = DateAdd("d", "-180", Date)
    
    Dim strSQL As String
    strSQL = "[Balance Due Date] <= #" & dt180minus & "#"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub
Private Sub cmdOpenReportAcctReceivable_Click()
    'DoCmd.OpenReport "Accounts Receivable", acViewReport, "", "([qryOutstandingARRPT1].[RetBal]<>0 and Orig_Atty='" & Me.cmbOrigAttyReport & "')"
    
    DoCmd.Close acReport, "Accounts Receivable", acSaveYes
    
    If Me.cmbOrigAtty <> "" Then
        DoCmd.OpenReport "Accounts Receivable", acViewPreview, "", "Orig_Atty='" & Me.cmbOrigAtty & "' And BalRetCalculated <> 0"
    Else
        DoCmd.OpenReport "Accounts Receivable", acViewPreview, "", "BalRetCalculated <> 0"
    End If
End Sub

Private Sub Command157_Click()

End Sub

Private Sub cmdRecordPDInvoice_Click()
    Dim sExistingReportName As String
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim strSQL As String

    
    If pcaempty(CaseID) Then
        MsgBox "Please select a client", , "TB CMS"
    Else
            
        'put the Invoice PDF copy to \Invoices folder

 
        sExistingReportName = "Invoice - Past Due"
        DocumentFileName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "PD Invoice" & " " & Date
        DocumentFileName = Replace(DocumentFileName, "/", "-")
        
        If Closed Then
            DocumentFolderName = GetClosedDocumentFolderName(CaseID, "Client Invoices")
        Else
            DocumentFolderName = GetDocumentFolderName(CaseID, "Client Invoices")
        End If
        
        AllInvoicesFolderName = GetAllInvoicesFolderName(CaseID)
        
    
        DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(CaseID, 0), acHidden  '& " AND [Past Due] = " & -1
        If FolderExistsCreate(DocumentFolderName, True) Then
            
            'save the invoice document to client document invoice folder
            DoCmd.OutputTo acOutputReport, sExistingReportName, acFormatPDF, DocumentFolderName & DocumentFileName & ".pdf"
            
            'save the document to _ALL INVOICES folder
            If FolderExistsCreate(AllInvoicesFolderName, True) Then
                DoCmd.OutputTo acOutputReport, sExistingReportName, acFormatPDF, AllInvoicesFolderName & DocumentFileName & ".pdf"
            End If
            
            DoCmd.Close acReport, sExistingReportName, acSaveNo
            
            If Not SaveCaseDocument(CaseID, "Client Invoices", DocumentFolderName & DocumentFileName & ".pdf") Then
                MsgBox "Failed to save Case Document record...", , "TB CMS"
            End If
            
            strSQL = "Insert into [tbl_InvoiceSent] "
                    strSQL = strSQL & "        (CaseID, InvSent, InvBalance)"
                    strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, " & Me.BalRetCalculated & ")"
            Debug.Print strSQL
            CurrentDb.Execute strSQL, dbSeeChanges
            Me.Requery
            
            MsgBox "Invoice Recorded and Saved", , "TB CMS"
            
        End If
    End If

End Sub

Private Sub Form_Load()
     DoCmd.ApplyFilter , "BalRetCalculated <> 0"
End Sub
Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub