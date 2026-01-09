' Component: Form_frmReceipt
' Type: document
' Lines: 45
' ============================================================

Option Compare Database

Private Sub cmdAddNew_Click()
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub cmdGenerateReceipt_Click()
    On Error GoTo ErrHandler_CaseID_Click
    If CurrentProject.AllReports("rptReceiptR").IsLoaded Then
        DoCmd.Close acReport, "rptReceiptR", acSaveNo
    End If
        'DoCmd.Close acForm, "frmReceipt"
        Me.Refresh
        DoCmd.OpenReport "rptReceiptR", acViewPreview, , "[ReceiptID]=" & Me.ReceiptID
     If CurrentProject.AllReports("rptReceiptR").IsLoaded Then
        DoCmd.Close acForm, "frmReceipt"
    End If
    
ErrHandler_CaseID_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description

'    Reports_rptReceiptUnbound.txtRDateR = txtRDate
'    Reports_rptReceiptUnbound.txtRFromR = txtRFrom
'    Reports_rptReceiptUnbound.txtRForR = txtRFor
'    Reports_rptReceiptUnbound.txtRMatterR = txtRMatter
'    Reports_rptReceiptUnbound.txtRDueR = txtRDue
'    Reports_rptReceiptUnbound.txtARBalanceR = txtARBalance
'    Reports_rptReceiptUnbound.txtRamountR = txtRAmount
'    Reports_rptReceiptUnbound.OptionCashR = OptionCash
'    Reports_rptReceiptUnbound.OptionCCR = OptionCC
'    Reports_rptReceiptUnbound.OptionCheckR = OptionCheck
    
End Sub

Private Sub Command26_Click()
    If CurrentProject.AllReports("rptReceiptR").IsLoaded Then
        DoCmd.Close acReport, "rptReceiptR", acSaveNo
    End If
'        DoCmd.Close acForm, "frmReceipt"
        DoCmd.OpenReport "rptReceiptR", acViewPreview, , "[ReceiptID]=" & Me.ReceiptID
End Sub

Private Sub Form_Load()
    DoCmd.GoToRecord , , acNewRec
End Sub