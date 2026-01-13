' Component: Form_frmInvoiceSent
' Type: document
' Lines: 140
' ============================================================

Option Compare Database

Private Sub cmdBillingOpenInvoiceFolder_Click()
Dim lCaseID As Variant
    
    lCaseID = Forms!frmClientLedger!CaseID
    
    If Not OpenDocumentFolder(lCaseID, "Client Invoices") Then
        MsgBox "Failed to open document folder...", , "TB CMS"
    End If
End Sub

Private Sub cmdRecondSentPDInvoice_Click()
    Dim sExistingReportName As String
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim CaseID As Variant
    Dim First_Name As Variant
    Dim Last_Name As Variant
    Dim strSQL As String
    
    CaseID = Forms!frmClientLedger!CaseID
    First_Name = Forms!frmClientLedger!First_Name
    Last_Name = Forms!frmClientLedger!Last_Name
    
    If pcaempty(CaseID) Then
        MsgBox "Please select a client", , "TB CMS"
    Else
            


        'put the Invoice PDF copy to \Invoices folder

 
        sExistingReportName = "Invoice - Past Due"
        DocumentFileName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "PD Invoice" & " " & Date
        DocumentFileName = Replace(DocumentFileName, "/", "-")
        
        If Form_frmClientLedger.Closed Then
            DocumentFolderName = GetClosedDocumentFolderName(Form_frmClientLedger.CaseID, "Client Invoices")
        Else
            DocumentFolderName = GetDocumentFolderName(Form_frmClientLedger.CaseID, "Client Invoices")
        End If
        
        AllInvoicesFolderName = GetAllInvoicesFolderName(Form_frmClientLedger.CaseID)
        
    
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
                    strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, " & Form_frmClientLedger.[Current Balance] & ")"
            Debug.Print strSQL
            CurrentDb.Execute strSQL, dbSeeChanges
            Me.Requery
        End If
    End If
End Sub

Private Sub cmdRecordSentInvoice_Click()
    Dim sExistingReportName As String
    Dim DocumentFileName As String
    Dim DocumentFolderName As String
    Dim AllInvoicesFolderName As String
    Dim CaseID As Variant
    Dim First_Name As Variant
    Dim Last_Name As Variant
    Dim strSQL As String
    Dim sFileFullPath As String
    
    
    CaseID = Forms!frmClientLedger!CaseID
    First_Name = Forms!frmClientLedger!First_Name
    Last_Name = Forms!frmClientLedger!Last_Name
    
    If pcaempty(CaseID) Then
        MsgBox "Please select a client", , "TB CMS"
    Else
            
        'put the Invoice PDF copy to \Invoices folder

 
        sExistingReportName = "Invoice2"
        DocumentFileName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Invoice" & " " & Date
        DocumentFileName = Replace(DocumentFileName, "/", "-")
        
        If Form_frmClientLedger.Closed Then
            DocumentFolderName = GetClosedDocumentFolderName(Form_frmClientLedger.CaseID, "Client Invoices")
        Else
            DocumentFolderName = GetDocumentFolderName(Form_frmClientLedger.CaseID, "Client Invoices")
        End If
        
        AllInvoicesFolderName = GetAllInvoicesFolderName(Form_frmClientLedger.CaseID)
        
        
           
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
                    strSQL = strSQL & " values (" & Form_frmClientLedger.CaseID & ", #" & Format(Date, "yyyy-MM-dd") & "#, " & Form_frmClientLedger.[Current Balance] & ")"
            Debug.Print strSQL
            CurrentDb.Execute strSQL, dbSeeChanges
            Me.Requery

        End If
    End If

End Sub
