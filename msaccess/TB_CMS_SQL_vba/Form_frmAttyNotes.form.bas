' Component: Form_frmAttyNotes
' Type: document
' Lines: 23
' ============================================================

Option Compare Database

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