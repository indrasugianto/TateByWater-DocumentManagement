' Component: Form_frmPersonalInjury
' Type: document
' Lines: 37
' ============================================================

Option Compare Database

Private Sub cmdPrintAdjLabel_Click()
Me.Refresh
 On Error GoTo ErrHandler_cmdPrintOCLabel_Click

    answer = MsgBox("Print Address Label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_adj_address_label").IsLoaded Then
            DoCmd.Close acReport, "rpt_adj_address_label", acSaveNo
        End If
        DoCmd.OpenReport "rpt_adj_address_label", acNormal, , "[CaseID]=" & CaseID
    End If
    
ErrHandler_cmdPrintOCLabel_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub CrStatusRep_Click()
  Dim sExistingReportName As String
    Dim sAttachmentName As String
 
    sExistingReportName = "rptPersInjuryStatus"
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