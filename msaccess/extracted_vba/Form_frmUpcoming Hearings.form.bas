' Component: Form_frmUpcoming Hearings
' Type: document
' Lines: 232
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

Private Sub Check44_AfterUpdate()

End Sub

Private Sub chkClientPresent_AfterUpdate()
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

Private Sub cmdHome_Click()
    DoCmd.openform "frmhome", acNormal
End Sub

Private Sub cmbHrgType_AfterUpdate()
    Call FilterMe
End Sub

Private Sub CmbOrigAtty_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmbPar_AfterUpdate()
    Call FilterMe
End Sub

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub

Sub FilterMe()
    Dim strSQL As String
    strSQL = "1=1"
    
    If Not IsNull(Me.cmbClients) Then strSQL = strSQL & " AND CaseID = " & Me.cmbClients
    If Not IsNull(Me.txtClient) Then strSQL = strSQL & " AND Name like '*" & Me.txtClient & "*'"
    If Not IsNull(Me.cmbHrgType) Then strSQL = strSQL & " AND HearingType = '" & Me.cmbHrgType & "'"
    If Not IsNull(Me.cmbOrigAtty) Then strSQL = strSQL & " AND Orig_Atty = '" & Me.cmbOrigAtty & "'"
    If Not IsNull(Me.cmbAssoc) Then strSQL = strSQL & " AND HandlingAtty_Case = '" & Me.cmbAssoc & "'"
    If Not IsNull(Me.cmbPar) Then strSQL = strSQL & " AND Paralegal = '" & Me.cmbPar & "'"
    If Not IsNull(Me.txtMatter) Then strSQL = strSQL & " AND Matter_type like '*" & Me.txtMatter & "*'"
    If Not IsNull(Me.txtNotes) Then strSQL = strSQL & " AND HrgResult like '*" & Me.txtNotes & "*'"
    If Not IsNull(Me.chkClientPresent) Then strSQL = strSQL & " AND ClientPresent = " & IIf(Me.chkClientPresent, 1, 0)
    'If Not IsNull(Me.chkPastDue) Then strSQL = strSQL & " AND [Past Due] = " & Me.chkPastDue
    'If Not IsNull(Me.chkNoBalance) Then strSQL = strSQL & " AND chkBalanceDue = " & Me.chkNoBalance
    
'    If Not IsNull(Me.chkNonZero) Then
'        If chkNonZero Then
'            strSQL = strSQL & " AND BalRet >0 "
'        Else
'            strSQL = strSQL & " AND BalRet =0 "
'        End If
'    End If
    
    Debug.Print strSQL
    
    Me.Filter = strSQL
    Me.FilterOn = True
End Sub

Sub FilterClear()

    Me.cmbClients = Null
    Me.cmbHrgType = Null
    Me.cmbOrigAtty = Null
    Me.cmbHrgType = Null
    Me.cmbPar = Null
    Me.txtMatter = Null
    Me.txtClient = Null
    Me.chkClientPresent = Null
    Me.txtNotes = Null
    
    Me.Filter = ""
    Me.FilterOn = False
End Sub


Private Sub cmdClearFilter_Click()
    Call FilterClear
End Sub

Private Sub cmdPrintNotes_Click()
    Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "rptClientNotes"
    sAttachmentName = Nz(First_Name) & " " & Nz(Last_Name) & " - " & "Notes" & " " & Date
  
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

Private Sub Combo40_BeforeUpdate(Cancel As Integer)
    If ClientPresent = 0 Then
        MsgBox "No need to send reminder of hearing where client not required to be present.", vbInformation, "TB CMS"
        Me.Undo
    End If
End Sub

Private Sub Image57_Click()
Application.FollowHyperlink Address:="https://app.zipwhip.com/messaging"
End Sub

Private Sub Text43_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
    Dim strValue As String
    strValue = Format(CCur(Me.txtOutstandingBalance), "Currency")
    Dim MediumTime As String
    MediumTime = Format(Me.HearingTime, "Medium Time")
    
    'Dim LValue As String
    'LValue = Format(Me.HearingTime, "Medium")
    'LValue = Format(Me.txtOutstandingBalance, "Currency")
     
    If Spanish = -1 And strValue > 0 Then
    Set OutlookApp = CreateObject("Outlook.Application")
    subject_ = "RECORDATORIO: Usted tiene corte el día " & Me.Hearing_Date & " a las " & MediumTime
    body_ = "Buenos días/tardes Sr./Sra. " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "Este es un recordatorio que su fecha de audiencia en la corte será el día " & Me.HearingDate & " a las " & Me.HearingTime & ". El saldo de su cuenta con nuestra oficina es " & strValue & ". Puede realizar el pago llamando a nuestra oficina al número a continuación con una tarjeta de débito/crédito o directamente por el internet en www.tatebywater.com/make-a-payment/." & vbCrLf & vbCrLf & vbCrLf & "Gracias. Si tiene alguna pregunta no dude en contactarnos."
    email_ = Me.Email
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    
    End With
    ElseIf Spanish = -1 And strValue = 0 Then
    Set OutlookApp = CreateObject("Outlook.Application")
    subject_ = "RECORDATORIO: Usted tiene corte el día " & Me.Hearing_Date & " a las " & MediumTime
    body_ = "Buenos días/tardes Sr./Sra.  " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "Este es un recordatorio que su fecha de audiencia en la corte será el día  " & Me.HearingDate & " a las " & Me.HearingTime & "." & vbCrLf & vbCrLf & vbCrLf & "Gracias. Si tiene alguna pregunta no dude en contactarnos."
    email_ = Me.Email
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
    ElseIf Spanish = 0 And strValue > 0 Then
    Set OutlookApp = CreateObject("Outlook.Application")
    subject_ = "REMINDER: You have court on " & Me.Hearing_Date & " at " & MediumTime
    body_ = "Good afternoon/morning, Mr./Mrs. " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "This is a reminder your hearing date will be on " & Me.HearingDate & " at " & Me.HearingTime & ". You have an outstanding balance with our office of " & strValue & ". You may tender your payment by calling our office at the number below with a debit / credit card or directly on our website at www.tatebywater.com/make-a-payment/." & vbCrLf & vbCrLf & vbCrLf & "Thank you and please let us know if you have any questions."
    email_ = Me.Email
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    
    End With
    ElseIf Spanish = 0 And strValue = 0 Then
    Set OutlookApp = CreateObject("Outlook.Application")
    subject_ = "REMINDER: You have court on " & Me.Hearing_Date & " at " & MediumTime
    body_ = "Good afternoon/morning, Mr./Mrs.  " & Me.First_Name & " " & Me.Last_Name & "," & vbCrLf & vbCrLf & "This is a reminder that your hearing date will be on  " & Me.HearingDate & " at " & Me.HearingTime & "." & vbCrLf & vbCrLf & vbCrLf & "Thank you and please let us know if you have any questions."
    email_ = Me.Email
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
    End If
End Sub

Private Sub txtClient_AfterUpdate()
    If Not IsNull(txtClient) Then Call FilterMe
End Sub

Private Sub txtMatter_AfterUpdate()
    If Not IsNull(txtMatter) Then Call FilterMe
End Sub

Private Sub txtNotes_AfterUpdate()
    If Not IsNull(txtNotes) Then Call FilterMe
End Sub