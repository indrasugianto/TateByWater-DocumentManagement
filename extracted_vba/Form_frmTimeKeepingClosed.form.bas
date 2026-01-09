' Component: Form_frmTimeKeepingClosed
' Type: document
' Lines: 557
' ============================================================

Option Compare Database

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
    
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '" & Me.txtClient & "*'"
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Private Sub ChkCloseTKFilter_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbAssoc_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbClients_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbclose_Click()
    DoCmd.Close acForm, Me.Name
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

Sub FilterClear()
    Me.cmbClients = Null
    Me.cmbOrigAtty = Null
    Me.cmbAssoc = Null
    Me.txtClient = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdCompCurr_Click()
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
    
'     If fncIsThereAZeroInMatterARForm = False Then
'        MsgBox "Please use the Full History Invoice.", , "TB CMS"
'        Exit Sub
'    End If
   
  
   
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
        
        MsgBox "Before printing make sure there is a fee.", , "TB CMS"
        DoCmd.OpenReport strReportName, acViewPreview, , "Bill_ID=" & Me.Bill_ID
        

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdCompFull_Click()
  Dim sAttachmentName As String
    sAttachmentName = IANumber & ": " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
 
 On Error GoTo ErrHandler_CaseID_Click
        
    If Me.InvoiceTotalAdvance = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceADV"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx"
    End If
    
    If Me.StatementLessTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceStmt"
    End If
    
    If Me.InvoiceNoAdvance = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKLessTrustRep"
    End If
    
     If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID
    Reports(sExistingReportName).Caption = sAttachmentName

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
   
End Sub

Private Sub cmdPreview_Click()

    On Error GoTo ErrHandler_CaseID_Click
    
    If Me.InvoiceTotalAdvance = -1 Then
        strReportName = "rpt_TKTotalAdvance"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
        strReportName = "rpt_TKExceedsTrust"
    End If
    
    If Me.StatementLessTrust = -1 Then
        strReportName = "rpt_TKLessTrust"
    End If
    
     If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
           
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
        
        'GH 2017-08-11 it was not compiling so I temporary disabled it
       DoCmd.OpenReport strReportName, acNormal, , "[Bill_ID]=" & Me.Bill_ID
    End If
    
ErrHandler_cmdPrintInvoice_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description

End Sub

Private Sub cmdPreview2_Click()
  On Error GoTo ErrHandler_CaseID_Click
'
'    If Me.TxtTrustBalance > 0 Then
'        strReportName = "Invoice Attach - Hourly w Discount"
'    Else
        strReportName = "Invoice Attach - Hourly"
'    End If

    'Debug.Print strReportName
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If

    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID

ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
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

Private Sub cmdRecordShort_Click()
    Dim strReportName As String
    Dim CaseID As Variant
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim strSQL As String
    
    CaseID = Me.CaseID
    
         
    If pcaempty(CaseID) Then
        MsgBox "Please select client.", , "TB CMS"
        Exit Sub
    End If
    
    DocumentFileName = IANumber & "- " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
    DocumentFileName = Replace(DocumentFileName, "/", "-")
        
    DocumentFolderName = GetDocumentFolderName(CaseID, "Client Invoices")
    AllInvoicesFolderName = GetAllInvoicesFolderName(CaseID)
    
 
    If Me.InvoiceTotalAdvance = -1 Then
      strReportName = "rpt_Comprehensive_InvoiceADVS"
    End If
    
    If Me.StatementLessTrust = -1 Then
      strReportName = "rpt_Comprehensive_InvoiceStmtS"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
      strReportName = "rpt_Comprehensive_InvoiceTKEx1S"
    End If
    
'    If Me.InvoiceAdvCostFee = -1 Then
'      strReportName = "rpt_Comprehensive_InvoiceTKEx2S"
'    End If
    
'    If Me.InvoiceCostHold = -1 Then
'      strReportName = "rpt_Comprehensive_InvoiceTKEx3CostsS"
'    End If

    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
     
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID, acHidden
    If FolderExistsCreate(DocumentFolderName, True) Then
    
        'save invoice to client document invoice folder
        DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, DocumentFolderName & DocumentFileName & ".pdf"
        
        'save invoice to _ALL INVOICES folder
        DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, AllInvoicesFolderName & DocumentFileName & ".pdf"
        
        DoCmd.Close acReport, strReportName, acSaveNo
    
        If Not SaveCaseDocument(CaseID, "Client Invoices", DocumentFolderName & DocumentFileName & ".pdf") Then
            MsgBox "Failed to save Case Document record...", , "TB CMS"
        End If
    End If
    
    
    strSQL = "Insert into [tbl_InvoiceSent] "
            strSQL = strSQL & "        (CaseID, TKDate, [TK Sent], TKNumber)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
    Debug.Print strSQL
    CurrentDb.Execute strSQL, dbSeeChanges


    strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
    Debug.Print strSQL
    CurrentDb.Execute strSQL, dbSeeChanges

    'Me.[Bill Sent].Requery
    Me.Requery
    
    MsgBox "Recorded and Saved", , "TB CMS"
    
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
  




'    strSQL = "Insert into [tbl_InvoiceSent] "
'            strSQL = strSQL & "        (CaseID, TKDate, [TK Sent], TKNumber)"
'            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL, dbSeeChanges
'
'
'    strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL, dbSeeChanges
'
'    'Me.[Bill Sent].Requery
'    Me.Requery
'    MsgBox "Recorded", , "TB CMS"

End Sub

Private Sub IANumber_Click()
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
    
    Form_frmClientLedger.tabControl.Pages(4).SetFocus
    [Form_Time Keeping].cmbBills = Me.txtBillID
    [Form_Time Keeping].cmbBills_AfterUpdate
    
ErrHandler_CaseNum_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdAddNewTK_Click()

    If CurrentProject.AllForms("frmClientLedger").IsLoaded Then
        DoCmd.Close acForm, "frmClientLedger", acSaveNo
        'DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbCases
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbClients, , , Me.CaseID
        Forms("frmClientLedger").SetFocus
    Else
'        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbCases
        DoCmd.openform "frmClientLedger", acNormal, , "[CaseID]=" & Me.cmbClients, , , Me.CaseID
        Forms("frmClientLedger").SetFocus
    End If
    
    Form_frmClientLedger.tabControl.Pages(4).SetFocus

'    strCaseID = Me.cmbCases
    strCaseID = Me.cmbClients
    
    If Not IsNull(strCaseID) Then
        currNr = Nz(DLookup("CountOfIANumber", "qry_get_time_keeping_numbers", "CaseID = " & strCaseID), 0)
        strIANumber = "TK-" & Nz(currNr) + 1
    End If
            
    strSQL = "Insert into [TB Time Keeping] (CaseID, [Bill Open], Discount, IANumber)"
    strSQL = strSQL & " values (" & strCaseID & ", #" & Date & "#, 0, '" & strIANumber & "')"
    
    Debug.Print "strSQL=" & strSQL
    CurrentDb.Execute strSQL
    
    Debug.Print "strIANumber=" & strIANumber
    
    [Form_Time Keeping].cmbBills.RowSource = "select * from qryBillList where CaseID=" & strCaseID
    [Form_Time Keeping].cmbBills.Requery
        
    maxBillID = DMax("Bill_ID", "qryBillList", "CaseID=" & strCaseID)
    [Form_Time Keeping].cmbBills = maxBillID
    Call [Form_Time Keeping].cmbBills_AfterUpdate
            
    [Form_Time Keeping].Requery
    '[Form_Time Keeping].Recordset.MoveLast
    
End Sub

Private Sub cmdRecord_Click()
    Dim strReportName As String
    Dim CaseID As Variant
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim strSQL As String
    
    CaseID = Me.CaseID
    
         
    If pcaempty(CaseID) Then
        MsgBox "Please select client.", , "TB CMS"
        Exit Sub
    End If
    
    DocumentFileName = IANumber & "- " & Nz(First_Name) & " " & Nz(Last_Name) & " - " & " " & [BilL Closed Date]
    DocumentFileName = Replace(DocumentFileName, "/", "-")
        
    DocumentFolderName = GetDocumentFolderName(CaseID, "Client Invoices")
    AllInvoicesFolderName = GetAllInvoicesFolderName(CaseID)
    
    
  
    If Me.InvoiceTotalAdvance = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceADV"
    End If
    
    If Me.StatementLessTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceStmt"
    End If
    
    If Me.InvoiceExceedsTrust = -1 Then
        strReportName = "rpt_Comprehensive_InvoiceTKEx1"
    End If
    
'    If Me.InvoiceAdvCostFee = -1 Then
'        strReportName = "rpt_Comprehensive_InvoiceTKEx2"
'    End If
    
'    If Me.InvoiceCostHold = -1 Then
'        strReportName = "rpt_Comprehensive_InvoiceTKEx3Costs"
'    End If
    
    If CurrentProject.AllReports(strReportName).IsLoaded Then
        DoCmd.Close acReport, strReportName, acSaveNo
    End If
     
    DoCmd.OpenReport strReportName, acViewPreview, , "[Bill_ID]=" & Me.Bill_ID, acHidden
    If FolderExistsCreate(DocumentFolderName, True) Then
        'save the invoice to client document invoice folder
        DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, DocumentFolderName & DocumentFileName & ".pdf"
        
        'save the invoice to _ALL INVOICES folder
        DoCmd.OutputTo acOutputReport, strReportName, acFormatPDF, AllInvoicesFolderName & DocumentFileName & ".pdf"
        
        DoCmd.Close acReport, strReportName, acSaveNo
    
        If Not SaveCaseDocument(CaseID, "Client Invoices", DocumentFolderName & DocumentFileName & ".pdf") Then
            MsgBox "Failed to save Case Document record...", , "TB CMS"
        End If
    
    End If
    
    
    strSQL = "Insert into [tbl_InvoiceSent] "
            strSQL = strSQL & "        (CaseID, TKDate, [TK Sent], TKNumber)"
            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
    Debug.Print strSQL
    CurrentDb.Execute strSQL, dbSeeChanges


    strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
    Debug.Print strSQL
    CurrentDb.Execute strSQL, dbSeeChanges

    'Me.[Bill Sent].Requery
    Me.Requery
      
      
    MsgBox "Recorded and Saved", , "TB CMS"
    
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
  




'    strSQL = "Insert into [tbl_InvoiceSent] "
'            strSQL = strSQL & "        (CaseID, TKDate, [TK Sent], TKNumber)"
'            strSQL = strSQL & " values (" & Me.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, -1, '" & Me.IANumber & "')"
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL, dbSeeChanges
'
'
'    strSQL = "Update [TB Time Keeping] set [Bill Sent] = #" & Format(Date, "yyyy-MM-dd") & "# where Bill_ID=" & Me.Bill_ID
'    Debug.Print strSQL
'    CurrentDb.Execute strSQL, dbSeeChanges
'
'    'Me.[Bill Sent].Requery
'    Me.Requery
'    MsgBox "Recorded", , "TB CMS"
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub