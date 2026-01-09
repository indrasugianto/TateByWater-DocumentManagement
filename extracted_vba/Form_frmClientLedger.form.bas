' Component: Form_frmClientLedger
' Type: document
' Lines: 1317
' ============================================================

Option Compare Database
'Private cls As New clsFormValidation


Private Sub CaseOpenDate_BeforeUpdate(Cancel As Integer)
If Not IsNull(Me.CaseOpenDate) Then
    answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
    If answer = "tb2740" Then
    Else
        Me.Undo
    End If
    End If
End Sub

Private Sub Closed_Click()
If IsNull(Me.Referral) Or IsNull(Me.Matter) Then
    Me.Closed = 0
    MsgBox "Matter and Client Source must be entered before closing." & vbCrLf & "If unknown, inquire before indicating unknown.", vbCritical, "TB CMS"
End If
End Sub

'Private Sub Closed_BeforeUpdate(Cancel As Integer)
'    If [Current Balance] <> 0 Or txtTrustBalance <> 0 Then
'        MsgBox "Client AR Balance and Trust Account Balance Must Both Be 0 to Close.", vbExclamation, "TB CMS"
'        Me.Undo
'    End If
'End Sub

'Private Sub Clsdate_BeforeUpdate(Cancel As Integer)
'    If Me.Closed = False Then
'        MsgBox "Please check Closed Box Before Entering Date.", vbExclamation, "TB CMS"
'        Me.Undo
'    End If
'End Sub

Private Sub cmbClients_GotFocus()
    txtLastClient = Null
End Sub

Private Sub cmbHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub

Private Sub cmdBillingOpenDocumentRetainer_Click()
    If Not OpenDocumentFile(Me.CaseID, "Retainer / Contract") Then
        MsgBox "Fail to open the document..."
    End If
End Sub

Private Sub cmdCaseList_Click()
    DoCmd.openform "frmCaseList", acNormal, , , , , Me.cmbYear.value
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "[Scan Location]"
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
'    Me.FilterOn = False
'    Me.Filter = "[Yr]= '" & Me.cmbYear & "'"
'    Me.FilterOn = True
    FilterMe
End Sub

Private Sub cmdClientReviewEmail_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Thank you for choosing TATE BYWATER"
    body_ = "Dear " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "It was our pleasure to assist you with your legal needs and we sincerely appreciate the trust you put in us at TATE BYWATER. If you have just 5 minutes, would you mind leaving an online review of your experience with us at the below links? We would greatly appreciate it as these reviews are an indispensable part of our business." & vbCrLf & vbCrLf & vbCrLf & "Thank you and please let us know if there is anything else we can help you with."
    email_ = Nz(Me.Email, "CLIENTEMAIL")
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
        .BCC = Me.Extended_Ledger
        
    End With
    strSQL = "Update [tblcase] set [ReviewReq] = #" & Format(Date, "yyyy-MM-dd") & "# where CaseID=" & Me.CaseID
    Debug.Print strSQL
    CurrentDb.Execute strSQL
End Sub

Private Sub cmdClientReviewEmailESP_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Gracias por elegir TATE BYWATER"
    body_ = "Estimado(a) " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "Fue un placer para nosotros asistirlo con sus necesidades legales y agradecemos sinceramente la confianza que depositó en TATE BYWATER. Si tiene 5 minutos, ¿nos haria el favor de dejar un testimonio de su experiencia con nosotros en los siguentes hiperenlaces? Lo agradeceríamos enormemente ya que estos testimonios son una parte indispensable de nuestro negocio." & vbCrLf & vbCrLf & vbCrLf & "Gracias.  Diganos si hay algo más con lo cual le podemos ayudar."
    email_ = Nz(Me.Email, "CLIENTEMAIL")
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
        .BCC = Me.Extended_Ledger
        
    End With
    strSQL = "Update [tblcase] set [ReviewReq] = #" & Format(Date, "yyyy-MM-dd") & "# where CaseID=" & Me.CaseID
    Debug.Print strSQL
    CurrentDb.Execute strSQL
End Sub

Private Sub cmdCloseCase_Click()
    If Me.Closed = -1 Then
        MsgBox "Case is already closed"
    Else
        If txtTrustBalance <> 0 Or [Current Balance] <> 0 Then
            MsgBox "Case Cannot be Closed if AR or Trust Balance is Not Zero", , "TB CMS"
        Else
            If IsNull(Me.Referral) Or IsNull(Me.Matter) Then
                Me.Closed = 0
                MsgBox "Matter and Client Source must be entered before closing." & vbCrLf & "If unknown, inquire before indicating Unknown.", vbCritical, "TB CMS"
            Else
                Me.Closed = -1
                Me.Clsdate = Date
                Me.Dirty = False
                
                If MsgBox("Do you want to move Client Common Drive Folder to CLOSED FILE SCANS?", vbYesNo, "TB CMS") = vbYes Then
                    If Not CopyDocumentToClosedFileScan(Me.CaseID) Then
                        MsgBox "Failed to copy the client folder to CLOSED FILE SCAN folder...", , "TB CMS"
                    Else
                        Me.Scanned = True
                    End If
                End If
                
                
                If MsgBox("Do you want to move Client Common Drive folder to the _CLOSED subfolder?", vbYesNo, "TB CMS") = vbYes Then
                    'move the client folder to closed subfolder
                    If Not MoveDocumentByCaseStatus(Me.CaseID, "Closed") Then
                        MsgBox "Failed to move client folder to closed folder...", , "TB CMS"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdClosingSheet_Click()
Dim sAttachmentName As String
    sAttachmentName = "Closing Sheet:" & Nz(First_Name) & " " & Nz(Last_Name)
    On Error GoTo ErrHandler_CmdClosingSheet
    DoCmd.Close acReport, "rpt_Main_Closing", acSaveYes
    DoCmd.OpenReport "rpt_Main_Closing", acViewPreview, , "[CaseID]=" & Me.CaseID
    Reports(sExistingReportName).Caption = sAttachmentName
ErrHandler_CmdClosingSheet:
    If Err.Number <> 0 Then ShowMessage Err.Description

End Sub

Private Sub cmdCreateFolder_Click()
    If Me.Closed Then
        MsgBox "Case is closed!", vbCritical, , "TB CMS"
    Else
        If pcaempty(Me.CaseID) Then
            MsgBox "Please select case...", , "TB CMS"
        Else
            If Not OpenDocumentFolder(Me.CaseID, "General") Then
                MsgBox "Failed to open folder...", , "TB CMS"
            End If
        End If
    End If
End Sub


Private Sub cmdCreateFolderSub_Click()
    If Me.DocumentType = "Init Intake, Notes, Documents" Or Me.DocumentType = "Client ID" Or Me.DocumentType = "Retainer / Contract" Or Me.DocumentType = "Closed Final" Or Me.DocumentType = "General" Then
        MsgBox "Sub folder cannot be created for this document type"
    Else
    If Not OpenDocumentFolder(Me.CaseID, DocumentType) Then
        MsgBox "Fail to open folder..."
    End If
    End If
End Sub

Private Sub cmdGenerateReviewEmail_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Thank you for choosing TATE BYWATER"
    body_ = "ATTENTION: Please forward the below email to the client ASAP: " & Nz(Me.Email, "CLIENTEMAIL") & ". BEFORE sending please review and make desired additions or corrections. REMEMBER to 1) copy client email into TO field 2) remove FW in subject line 2) remove forwarding address block 3)sign email and 4) bcc Lizzie and Karen)." & vbCrLf & vbCrLf & "Dear " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "It was my pleasure to assist you with your legal needs and I sincerely appreciate the trust you put in me and the staff at TATE BYWATER. If you have just 5 minutes, would you mind leaving an online review of your experience with us at the below links? I would greatly appreciate it as these reviews are an indispensable part of our business." & vbCrLf & vbCrLf & vbCrLf & "Thank you and please let me know if there is anything else I can help you with."
    email_ = Me.Extended_Ledger
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
        
    End With
End Sub

Private Sub cmdGenerateReviewEmailESP_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Gracias por elegir TATE BYWATER"
    body_ = "ATTENTION: Please forward the below email to the client ASAP: " & Nz(Me.Email, "CLIENTEMAIL") & ". BEFORE sending please review and make desired additions or corrections. REMEMBER to 1) copy client email into TO field 2) remove FW in subject line 2) remove forwarding address block 3)sign email and 4) bcc Lizzie and Karen)." & vbCrLf & vbCrLf & "Estimado(a) " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "Fue un placer asistirlo con sus necesidades legales y agradezco sinceramente la confianza que depositó en mi y TATE BYWATER. Si tiene 5 minutos, ¿me haria el favor de dejar un testimonio de su experiencia con nosotros en los siguentes hiperenlaces? Lo agradeceríamos enormemente ya que estos testimonios son una parte indispensable de nuestro negocio." & vbCrLf & vbCrLf & vbCrLf & "Gracias.  Digame si hay algo más con lo cual le puedo ayudar."
    email_ = Me.Extended_Ledger
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
        
    End With
End Sub

Private Sub cmdMerge_Click()
    Set objWord = GetObject("S:\Merge docs\EOA - JDR - Master.docx", "Word.Document")
    ' Make Word visible.
    objWord.Application.Visible = True
    ' Set the mail merge data source as the CMS database.
    objWord.MailMerge.OpenDataSource _
    Name:="s:\CMS\TB CMS BE 1.2.accdb", _
    LinkToSource:=True, _
    Connection:="QUERY qryMergeTest", _
    SQLStatement:="SELECT * FROM [qryMergeTest]"
    ' SQLStatement:="SELECT Last_Name, First_Name, Organization, Address_1,City, State, Zip FROM [MailMergeEvent]"
    
    ' Execute the mail merge.
    objWord.MailMerge.Execute
End Sub

Private Sub cmdOpenClosedFinal_Click()
    If Not OpenDocumentFile(Me.CaseID, "Closed Final") Then
        MsgBox "Failed to open the document...", , "TB CMS"
    End If
End Sub

Private Sub cmdOpenDocumentClientID_Click()
    If Not OpenDocumentFile(Me.CaseID, "Client ID") Then
        MsgBox "Failed to open the document...", , "TB CMS"
    End If
End Sub

Private Sub cmdOpenDocumentFolderCorrespondence_Click()
    If Not OpenDocumentFolder(Me.CaseID, "Correspondence: Letters and Emails") Then
        MsgBox "Failed to open folder...", , "TB CMS"
    End If
End Sub

Private Sub cmdOpenDocumentFolderFinance_Click()
    If Not OpenDocumentFolder(Me.CaseID, "Discovery") Then
        MsgBox "Failed to open folder...", , "TB CMS"
    End If
End Sub

Private Sub cmdOpenDocumentFolderFull_Click()
    If Not OpenDocumentFolder(Me.CaseID, "General") Then
        MsgBox "Failed to open folder...", , "TB CMS"
    End If
End Sub
Private Sub cmdOpenDocumentFolderInvoices_Click()
    If Not OpenDocumentFolder(Me.CaseID, "Client Invoices") Then
        MsgBox "Failed to open folder...", , "TB CMS"
    End If
End Sub


Private Sub cmdOpenInitialIntake_Click()
    If Not OpenDocumentFile(Me.CaseID, "Init Intake, Notes, Documents") Then
        MsgBox "Failed to open the document...", , "TB CMS"
    End If
End Sub

Private Sub cmdOpenRetainer_Click()
    If Not OpenDocumentFile(Me.CaseID, "Retainer / Contract") Then
        MsgBox "Failed to open the document...", , "TB CMS"
    End If
End Sub

Private Sub cmdPrintFileLabel_Click()
Me.Refresh
 On Error GoTo ErrHandler_cmdPrintFileLabel_Click

    answer = MsgBox("Print File Folder Label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_File_Folder_Label").IsLoaded Then
            DoCmd.Close acReport, "rpt_File_Folder_Label", acSaveNo
        End If
        DoCmd.OpenReport "rpt_File_Folder_Label", acNormal, , "[CaseID]=" & CaseID
    End If
    
ErrHandler_cmdPrintFileLabel_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdPrintLabelP_Click()
    Me.Refresh
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

Private Sub cmdPrintOCLabel_Click()
Me.Refresh
 On Error GoTo ErrHandler_cmdPrintOCLabel_Click

    answer = MsgBox("Print Address Label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_opp_counsel_address_label").IsLoaded Then
            DoCmd.Close acReport, "rpt_opp_counsel_address_label", acSaveNo
        End If
        DoCmd.OpenReport "rpt_opp_counsel_address_label", acNormal, , "[CaseID]=" & CaseID
    End If
    
ErrHandler_cmdPrintOCLabel_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdReopenCase_Click()
    If (Me.Closed = True) Then
        Me.Closed = False
        Me.Dirty = False
        Clsdate = ""
    
        If MsgBox("Do you want to move back the Client folder?", vbYesNo) = vbYes Then
            'move back the Client folder
            If Not MoveDocumentByCaseStatus(Me.CaseID, "Open") Then
                MsgBox "Failed to move the folder back to open...", , "TB CMS"
            End If
        End If
    End If
End Sub

Private Sub cmdScan_Click()
Dim ScannerDirectory As String
Dim SelectedFileName As String
Dim CaseStatus As String
    If Me.Closed Then
        MsgBox "Case is closed!", vbCritical, "TB CMS"
    Else
        If pcaempty(Me.CaseID) Then
            MsgBox "Please select the case before continue...", , "TB CMS"
        Else
            If pcaempty(DocumentType) Then
                MsgBox "Please select the type of document that you want to scan/save", , "TB CMS"
            Else
                'select the scanned file
                ScannerDirectory = GetScannerFolder()
                SelectedFileName = SelectFileDialog("Select scanned file", ScannerDirectory, "")
                
                If Not pcaempty(SelectedFileName) Then
                    'if the document type is Closed Final, need to put copy on the CLOSED FINAL SCAN folder as well
                    If Me.Closed Then
                        CaseStatus = "Closed"
                    Else
                        CaseStatus = "Open"
                    End If
                    If SaveScannedFileAs(Me.CaseID, DocumentType, SelectedFileName, CaseStatus) Then
                        If DocumentType = "Closed Final" Then
                            Me.Scanned = True
                        End If
                        MsgBox "Scanned file is succesfully stored and assigned to this case.", , "TB CMS"
                    End If
    
                End If
            End If
        End If
    End If
End Sub

'Private Sub Command128_Click()
'    lngChnc = 1
'    DoCmd.RunCommand acCmdSaveRecord '=== TO SAVE ANY UNSAVED NEW CASE.
'    DoCmd.openform "frmFamilyLaw", acNormal, , "CaseID=" & Me.CaseID, , acWindowNormal, Me.CaseID
'End Sub

Private Sub cmdStatementTrustAcct_Click()
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

Private Sub Command1085_Click()

End Sub

Private Sub Command912_Click()
    If Not IsNull(Me.Ocounsel) Then
    Dim oApp As Outlook.Application
    Dim oMail As MailItem
    Set oApp = CreateObject("Outlook.application")
    'oApp.Visible = True
    Set oMail = oApp.CreateItem(olMailItem)
    
    oMail.Body = ""
    oMail.Subject = ""
    oMail.To = Me.OC_Email
    oMail.Display
    'oMail.Send
    Set oMail = Nothing
    Set oApp = Nothing
    Else
    MsgBox "Please enter opposing counsel email address.", , "TB CMS"
    End If
End Sub

Private Sub CrStatusRep_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "rptCriminalStatus"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Status Report" & " " & Date
  
    On Error GoTo ErrHandler_cmdInvoice

    DoCmd.Close acReport, sExistingReportName, acSaveYes
    DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
    Reports(sExistingReportName).Caption = sAttachmentName
ErrHandler_cmdInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
End Sub

Private Sub Email_BeforeUpdate(Cancel As Integer)
    Dim ctlval As Control
    Set ctlval = Form.ActiveControl
    Cancel = EmailCheck(ctlval)
    Set ctlval = Nothing
End Sub


Private Sub Form_Open(Cancel As Integer)

    If Not IsNull(Me.OpenArgs) Then
        Me.cmbClients = Me.OpenArgs
    Else
        Me.cmbClients = Null
    End If
    
    'Me.chkClosed = Null
    Me.cmbFileNumbers = Null
    Call FilterMe

End Sub

Private Sub Form_Timer()
'    Me![txttenths].Visible = False    'make the Text Box invisible after 5 secs
'    Me.TimerInterval = 0              'disable the Timer
End Sub



Private Sub Image57_Click()
Application.FollowHyperlink Address:="https://app.zipwhip.com/messaging"
End Sub

Private Sub Image830_Click()
DoCmd.openform "frmhome", acNormal
End Sub

Private Sub cmdEmailPastDue_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "Invoice - Past Due"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Past Due Invoice" & " " & Date
 
'    If Forms!Billing![Balance Due Date] > Date Then
'    MsgBox "This Invoice is Not Past Due.", , "TB CMS"
    On Error GoTo ErrHandler_cmdPastDueInvoice
    If Nz(Me.CaseID, 0) = 0 Then
        ShowMessage "Please select a client."
    Else
        DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0) '& " AND [Past Due] = " & -1
        Reports(sExistingReportName).Caption = sAttachmentName
        DoCmd.SendObject acSendReport, , acFormatPDF, Nz(Me.Email), , , Nz(First_Name) & " " & Nz(Last_Name) & ": " & "Past Due Invoice", Nz(Title) & " " & Nz(Last_Name) & ": " & vbLf & vbLf & "Please see the attached invoice. To make payment please use our payment portal at www.tatebywater.com/make-a-payment or send a check/money order payable to Tate Bywater or call the office at (703) 938-5100. We appreciate your prompt payment." & vbLf & vbLf & "Thank you.", True
    
    End If
ErrHandler_cmdPastDueInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
'End If
End Sub
Private Sub cmdEmailNoBalance_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "Invoice - No Balance Due"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "No Balance Invoice" & " " & Date
    
    
    If Me.Current_Balance <> 0 Then
        MsgBox "Client has an Account Receivable Balance", , "TB CMS"
    Else
        DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
        Reports(sExistingReportName).Caption = sAttachmentName
        DoCmd.SendObject acSendReport, , acFormatPDF, Nz(Me.Email), , , Nz(First_Name) & " " & Nz(Last_Name) & ": " & "Invoice (Zero Balance)", Nz(Title) & " " & Nz(Last_Name) & ": " & vbLf & vbLf & "Please see the attached invoice. You currently have no balance due on your account." & vbLf & vbLf & "Thank you.", True, "No_Balance_Invoice" & Date
    
    End If
ErrHandler_cmdPastDueInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
'End If
End Sub
Private Sub cmdEmailFullHistory_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "Invoice2"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Invoice" & " " & Date
    
    On Error GoTo ErrHandler_cmdPastDueInvoice
    If Nz(Me.CaseID, 0) = 0 Then
        ShowMessage "Please select a client."
    Else
        DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0) '& " AND [Past Due] = " & -1
        Reports(sExistingReportName).Caption = sAttachmentName
        DoCmd.SendObject acSendReport, , acFormatPDF, Nz(Me.Email), , , Nz(First_Name) & " " & Nz(Last_Name) & ": " & "Invoice", Nz(Title) & " " & Nz(Last_Name) & ": " & vbLf & vbLf & "Please see the attached invoice. To make payment please use our payment portal at www.tatebywater.com/make-a-payment or send a check/money order payable to Tate Bywater or call the office at (703) 938-5100. We appreciate your prompt payment." & vbLf & vbLf & "Thank you.", True, "Past_Due_Invoice" & Date
    
    End If
ErrHandler_cmdPastDueInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
'End If
End Sub

Private Sub Number__BeforeUpdate(Cancel As Integer)
    answer = InputBox("Please input the Admin Password", "TB CMS: Admin Pass Required")
    If answer = "tb2740" Then
    Else
        Me.Undo
    End If
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



Private Sub chkClosed_AfterUpdate()
    Call FilterMe
End Sub

Sub FilterMe()
    Dim strSQL As String
    Dim strWhere As String
    strSQL = Me.RecordSource
    
    
    
    strWhere = ""
        
    If Not IsNull(Me.chkClosed) Then strWhere = strWhere & " AND closed = " & IIf(Me.chkClosed, 1, 0)
    If Not IsNull(Me.cmbClients) Then strWhere = strWhere & " AND CaseID= " & Me.cmbClients
    If Not IsNull(Me.cmbFileNumbers) Then strWhere = strWhere & " AND  CaseID = " & Me.cmbFileNumbers
    If Not IsNull(Me.cmbYear) Then strWhere = strWhere & " AND [Yr] = '" & Me.cmbYear & "'"
    
    
    If Len(strWhere) = 0 Then strWhere = " AND 1 = 0"
    
    strWhere = Mid(strWhere, 6, Len(strWhere))
    
    strSQL = pcaPutSQLSection(strSQL, strWhere, "WHERE")
    
    Me.RecordSource = strSQL
    'Me.Requery
    
    
    'Debug.Print strSQL
    
End Sub
Sub FilterClear()
    'clear controls:
    Me.cmbYear = Null
    Me.chkClosed = Null
    Me.cmbClients = Null
    Me.cmbFileNumbers = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub

Private Sub Refresh_Click()
    Me.frmMatter.Requery
    Me.frmTrustAccount.Requery
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

'Private Sub cmdSave_Click()
'    'gaz 2017-07-23 disabled:
'
''    If Me.Dirty Then
''        cls.ExeCommand SaveRec 'gh: weird validation?
''    End If
''    CaseGenerator
'
'    'gaz 2017-07-23
'
'End Sub

Private Sub cmdAddNew_Click()
    'cls.ExeCommand Addrec
    Dim strSQL As String
    Dim strWhere As String
    
    strSQL = Me.RecordSource
    
    strWhere = "1=1"
    
    strSQL = pcaPutSQLSection(strSQL, strWhere, "WHERE")
    
    Me.RecordSource = strSQL
    
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

'Private Sub cmdStatemenTrustAccount_Click()
'    On Error GoTo ErrHanlder_cmdStatementrustAccount
'        StatementofTrustAccount_Filter = True
'        DoCmd.OpenReport "Statement of Trust Account", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
'ErrHanlder_cmdStatementrustAccount:
'    If Err.Number = 2501 Then
'        ShowMessage "No records found!"
'    ElseIf Err.Number <> 0 Then
'        ShowMessage Err.Description
'    End If
'
'End Sub

'Private Sub Command225_Click()
'On Error Resume Next
'    DoCmd.RunCommand acCmdRecordsGoToNew
'End Sub

'gaz 2017-05-05
Private Sub cmdPastDueInvoice_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "Invoice - Past Due"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Past Due Invoice" & " " & Date



'    If Forms!Billing![Balance Due Date] > Date Then
'    MsgBox "This Invoice is Not Past Due.", , "TB CMS"
    On Error GoTo ErrHandler_cmdPastDueInvoice
    If Nz(Me.CaseID, 0) = 0 Then
        ShowMessage "Please select a client."
    Else
        DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0) '& " AND [Past Due] = " & -1
        Reports(sExistingReportName).Caption = sAttachmentName
    End If
ErrHandler_cmdPastDueInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
'End If
End Sub

'gaz 2017-05-05
Private Sub cmdBalanceInvoice_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "Invoice - No Balance Due"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "No Balance Invoice" & " " & Date
    
    On Error GoTo ErrHandler_cmdBalanceInvoice
If Me.Current_Balance <> 0 Then
        MsgBox "Client has an Account Receivable Balance", , "TB CMS"
    Else
        DoCmd.Close acReport, sExistingReportName, acSaveNo
        DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
        Reports(sExistingReportName).Caption = sAttachmentName
    End If
   '
      
    'DoCmd.Close acReport, "Invoice - No Balance Due", acSaveYes
    'DoCmd.OpenReport "Invoice - No Balance Due", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0) & " AND [chkBalanceDue] = " & 0
    'DoCmd.OpenReport "Invoice - No Balance Due", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
        
ErrHandler_cmdBalanceInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
    
End Sub

'gaz 2017-05-05
'Private Sub CmdFullHistoryInvoice_Click()
'
'    On Error GoTo ErrHandler_cmdInvoice
'
'    DoCmd.Close acReport, "Invoice", acSaveYes
'    DoCmd.OpenReport "Invoice", acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
'ErrHandler_cmdInvoice:
'    If Err.Number = 2501 Then
'        ShowMessage "No records found!"
'    ElseIf Err.Number <> 0 Then
'        ShowMessage Err.Description
'    End If
'End Sub

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

Private Sub cmdPrintLabel_Click()
    Me.Refresh
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

'Private Sub cmdSubmitToFamilyLaw_Click()
''    Dim strSQL As String
''    strSQL = "Insert into [Family Law - Divorce]  ([CaseID]) values (" & lngCaseID & ")"
''    DoCmd.SetWarnings False
''    DoCmd.RunSQL (strSQL)
'    lngChnc = 1
'      DoCmd.RunCommand acCmdSaveRecord '=== TO SAVE ANY UNSAVED NEW CASE.
'    DoCmd.openform "frmFamilyLaw", acNormal, , "CaseID=" & lngCaseID, , acWindowNormal, lngCaseID
'End Sub



'Private Sub Case_Letter_AfterUpdate()
    'Me.yr = Format(Date, "yy") 'Right(Year(Date), 2)
    'Me.Year = Format(Date, "yy") 'Right(Year(Date), 2)
'End Sub


'Private Sub Last_Name_AfterUpdate()
    '?
'End Sub

Private Sub cmdClose_Click()
    'cls.ExeCommand Cancelrec
    DoCmd.Close acForm, Me.Name, acSave
End Sub

Private Sub cmbClients_AfterUpdate()
    cmbFileNumbers = Null
    Call FilterMe
End Sub

Private Sub cmbFileNumbers_AfterUpdate()
    cmbClients = Null
    Call FilterMe
End Sub



'Private Sub cmdHourlyBillList_Click()
'    'SHOW TK BILLS FORM
'
'    If IsNull(CaseID) Then
'        MsgBox "Please select a client"
'        Exit Sub
'    End If
'    DoCmd.openform "Time Keeping", acNormal, , "[CaseID]=" & Nz(Me.CaseID, 0), , , "Hide Filter"
'
'    'GH 2017-08-11 it was not compiling so I temporary disabled it
'    [Form_Time Keeping].cmbClients = Me.CaseID
'
'
'    If Not IsNull(CaseID) Then
'        Call [Form_Time Keeping].cmbClients_AfterUpdate
'    End If
'End Sub

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



'GH 2018-08-12 Port this code to the TK form?????
'Private Sub AddNewTK_Click()
'    strCaseID = Form_frmClientLedger.CaseID
'
'    If Not IsNull(strCaseID) Then
'        currNr = DLookup("CountOfIANumber", "qry_get_time_keeping_numbers", "CaseID = " & strCaseID)
'        strIANumber = "TK-" & Nz(currNr) + 1
'    End If
'
'    strSQL = "Insert into [TB Time Keeping] (CaseID, [Bill Sent], Discount, IANumber)"
'    strSQL = strSQL & " values (" & strCaseID & ", #" & Date & "#, 0, '" & strIANumber & "')"
'
'    Debug.Print "strSQL=" & strSQL
'    CurrentDb.Execute strSQL
'
'    Debug.Print "strIANumber=" & strIANumber
'
'
'    'GH 2017-08-11 it was not compiling so I temporary disabled it
'    'Me.cmbBills.RowSource = "select * from qryBillList where CaseID=" & strCaseID
'    'Me.cmbBills.Requery
'    'GH 2017-08-11 it was not compiling so I temporary disabled it
'
'
'
'    'Me.Requery
'    'Me.Recordset.MoveLast
'    'Me.filter = "[CaseID]=" & strCaseID & " and IANumber = '" & strIANumber & "'"
'    'Me.FilterOn = True
'
'    maxBillID = DMax("Bill_ID", "qryBillList", "CaseID=" & strCaseID)
'
'    'GH 2017-08-11 it was not compiling so I temporary disabled it
'    'Me.cmbBills = maxBillID
'    'Call cmbBills_AfterUpdate
'    'GH 2017-08-11 it was not compiling so I temporary disabled it
'
'
'
'    'On Error GoTo errorhandler
'    'If IsNull(IANumber) Then IANumber = strIANumber
'
'    'also set the Discount to 0
'    'If IsNull(Discount) Then Discount = 0
'
'    '[Form_Time Keeping].Requery
'    '[Form_Time Keeping].Recordset.MoveLast
'    'Me.cmbBills.Requery
'
'    'strCaseID = CaseID
'    'DoCmd.GoToRecord , , acNewRec
'    'CaseID = strCaseID
'    'Call setIANumber
'    'Me.Bill_Sent = Date
'
'    'Me.filter = "[CaseID]=" & strCaseID
'    'Me.FilterOn = True
'
''    Exit Sub
''errorhandler:
''    If Err.Description <> "Cannot add record(s); join key of table 'TB Time Keeping' not in recordset." Then
''        MsgBox Err.Description
''    End If
'End Sub
'
'Private Sub cmdCreateHourlyBill_Click()
'    'NEW TK BILL
'    'On Error GoTo ErrHandler_cmdCreateHourlyBill
'
'    'Check if this CaseID is in the Time Keeping table:
''    Dim rs As DAO.Recordset
''    If DCount("CaseID", "TB Time Keeping", "CaseID=" & Nz(Me.CaseID, 0)) = 0 Then
''        Set rs = CurrentDb.OpenRecordset("TB Time Keeping")
''        rs.AddNew
''        rs.Fields("CaseID") = Nz(Me.CaseID, 0)
''        rs.Update
''        rs.Close
''        Set rs = Nothing
''    Else
''        'do  nothing
''    End If
'
'    DoCmd.openform "Time Keeping", acNormal, , "[CaseID]=" & Nz(Me.CaseID, 0), , , "Hide Filter"
'    Call [Form_Time Keeping].cmdAddNew_Click
'    '[Form_Time Keeping].Recordset.MoveLast
'
'    '[Form_Time Keeping].cmbClients = Me.CaseID
'    'Call [Form_Time Keeping].cmbClients_AfterUpdate
'
''ErrHandler_cmdCreateHourlyBill:
''    If Err.Number <> 0 Then MsgBox Err.Description
'End Sub

Private Sub OrigAtty_AfterUpdate()
    Me.FileNo.Requery
    Me.txtFileNo.Requery
    If Me.Orig_Atty = "JRT" Then
        Me.Extended_Ledger.value = "jtate@tatebywater.com"
    End If
    If Me.Orig_Atty = "DEB" Then
        Me.Extended_Ledger.value = "dbywater@tatebywater.com"
    End If
    If Me.Orig_Atty = "GBF" Then
        Me.Extended_Ledger.value = "gfuller@tatebywater.com"
    End If
    If Me.Orig_Atty = "PM" Then
        Me.Extended_Ledger.value = "pmickelsen@tatebywater.com"
    End If
    If Me.Orig_Atty = "TDT" Then
        Me.Extended_Ledger.value = "ttull@tatebywater.com"
    End If
    If Me.Orig_Atty = "MK" Then
        Me.Extended_Ledger.value = "mkennedy@tatebywater.com"
    End If
    If Me.Orig_Atty = "MT" Then
        Me.Extended_Ledger.value = "mtaylor@tatebywater.com"
    End If
    If Me.Orig_Atty = "KDB" Then
        Me.Extended_Ledger.value = "kbigus@tatebywater.com"
    End If
    If Me.Orig_Atty = "CMH" Then
        Me.Extended_Ledger.value = "cmhiggins@tatebywater.com"
    End If
    If Me.Orig_Atty = "RLF" Then
        Me.Extended_Ledger.value = "rfredericks@tatebywater.com"
    End If
    If Me.Orig_Atty = "NH" Then
        Me.Extended_Ledger.value = "nhasan@tatebywater.com"
    End If
    If Me.Orig_Atty = "BA" Then
        Me.Extended_Ledger.value = "bader@tatebywater.com"
    End If
     If Me.Orig_Atty = "RRL" Then
        Me.Extended_Ledger.value = "rlucero@tatebywater.com"
    End If
End Sub

Private Sub CaseOpenDate_AfterUpdate()
    
    'Me.yr = Format(Date, "yy") 'Right(Year(Date), 2)
    year2digits = Format(CaseOpenDate, "yy")
    Me.yr = year2digits
    nrMax = Nz(DMax("Number_", "tblCase", "Yr='" & Me.yr & "'"), 0)
    'nrMax = Nz(DMax("Number_", "tblCase", "Yr='" & year2digits & "'"), 0)
    Debug.Print "nrMax: " & nrMax
    Me.Number_ = nrMax + 1
End Sub

Private Sub Form_Load()
'    Me.filter = "[Yr]= '" & Right(Year(Now), 2) & "' and Closed=0"
'    Me.FilterOn = True
'
'    'add new record:
'    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub Command353_Click()
    Dim sAttachmentName As String
    sAttachmentName = "Closing Sheet:" & Nz(First_Name) & " " & Nz(Last_Name)
    On Error GoTo ErrHandler_CmdClosingSheet
    DoCmd.Close acReport, "rpt_Main_Closing", acSaveYes
    DoCmd.OpenReport "rpt_Main_Closing", acViewPreview, , "[CaseID]=" & Me.CaseID
    Reports(sExistingReportName).Caption = sAttachmentName
ErrHandler_CmdClosingSheet:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub Retainer_AfterUpdate()
    Me.frmMatter.Requery
    Me.Refresh
End Sub

Private Sub cmdOutlookTask_Click()
    If Not IsNull(Me.Email) Then
    Dim oApp As Outlook.Application
    Dim oMail As MailItem
    Set oApp = CreateObject("Outlook.application")
    'oApp.Visible = True
    Set oMail = oApp.CreateItem(olMailItem)
    
    oMail.Body = ""
    oMail.Subject = ""
    oMail.To = Me.Email
    oMail.Display
    'oMail.Send
    Set oMail = Nothing
    Set oApp = Nothing
    Else
    MsgBox "Please enter client email address.", , "TB CMS"
    End If
End Sub
Private Sub Form_Current()

    lngCaseID = Nz(Me.CaseID, 0)

    Debug.Print "CaseID: " & GetCaseID
    Me.cmbClients = Me.CaseID
    Me.cmbFileNumbers = Me.CaseID

'    If DLookup("Balance", "qryMatter", "CaseID=" & GetCaseID()) > 0 Then
'        Me.frmBilling.Form.Controls("chkBalanceDue") = True
'    Else
'        Me.frmBilling.Form.Controls("chkBalanceDue") = False
'    End If

    If IsNull(Form_frmMatter.OrderNr) Then
        If Not IsNull(Form_frmClientLedger.CaseID) Then
            Call Form_frmMatter.reorderByDateMatter(Form_frmClientLedger.CaseID)
            Form_frmMatter.Requery
        End If
    End If

    If IsNull(Form_frmTrustAccount.OrderNr) Then
        Call Form_frmTrustAccount.reorderByDateTrustAccount
        Form_frmTrustAccount.Requery
    End If
    
End Sub

Private Sub CommandFullHistoryInvoice_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "Invoice2"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Invoice" & " " & Date
  
    On Error GoTo ErrHandler_cmdInvoice

    DoCmd.Close acReport, sExistingReportName, acSaveYes
    DoCmd.OpenReport sExistingReportName, acViewPreview, , "[CaseID]=" & Nz(Me.CaseID, 0)
    Reports(sExistingReportName).Caption = sAttachmentName
ErrHandler_cmdInvoice:
    If Err.Number = 2501 Then
        ShowMessage "No records found!"
    ElseIf Err.Number <> 0 Then
        ShowMessage Err.Description
    End If
End Sub

Private Sub cmdInvoice_Click()
        
    If fncIsThereAZeroInMatterARForm = False Then
        MsgBox "Please use the Full History Invoice.", , "TB CMS"
        Exit Sub
    End If
        
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
    
    DoCmd.OpenReport "New Invoice", acViewPreview, , strFilter
End Sub

Function fncIsThereAZeroInMatterARForm() As Boolean
    Dim rs As Recordset
    Set rs = Form_frmMatter.Recordset
    Do Until rs.EOF
        bal = fncGetMatterARBalanceWithCaseID(rs("OrderNr"), Form_frmClientLedger.CaseID)
        Debug.Print rs("OrderNr") & "=" & bal
        If bal = 0 Then
            fncIsThereAZeroInMatterARForm = True
            Exit Function
        End If
        rs.MoveNext
    Loop
    fncIsThereAZeroInMatterARForm = False
End Function