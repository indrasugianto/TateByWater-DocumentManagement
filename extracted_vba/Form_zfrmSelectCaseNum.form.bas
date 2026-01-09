' Component: Form_zfrmSelectCaseNum
' Type: document
' Lines: 25
' ============================================================

Option Compare Database

Private Sub CmdOpenInvoiceAttachReport_Click()
On Error GoTo ErrHnd
Me.Visible = False
If Nz(Me.cboCaseNum, "") <> "" Then
    DoCmd.OpenReport "Invoice Attach - Hourly ", acViewPreview, , "CaseID=" & Me.cboCaseNum, acWindowNormal

Else
    DoCmd.OpenReport "Invoice Attach - Hourly ", acViewPreview, , , acWindowNormal
End If
    'DoCmd.Close acForm, "frmSelectCaseNum", acSaveNo

Exit Sub
ErrHnd:
    If Err.Number = 2501 Then
        ShowMessage "No records found"
    ElseIf Err.Number <> 0 Then
        ErrMsg "CmdOpenInvoiceAttachReport"
    End If
End Sub



