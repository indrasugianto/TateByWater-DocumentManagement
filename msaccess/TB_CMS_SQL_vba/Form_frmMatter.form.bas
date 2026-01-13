' Component: Form_frmMatter
' Type: document
' Lines: 273
' ============================================================

Option Compare Database

'Dim OldDate 'As Date
'Dim OldMatterID As Long
'Dim OldPay_Outlay As String
'Dim OldPayment As String
'Dim OldCharge As Currency
'Dim OldFirmPrepaid As Boolean

'Private Sub Charge_AfterUpdate()
'    Text39.Requery
'End Sub
'
'Private Sub Payment_AfterUpdate()
'    Text39.Requery
'End Sub

Private Sub Charge_AfterUpdate()
    'Me.Recalc
    If Me.Dirty Then Me.Dirty = False
    'Me.Payment.SetFocus
    If IsNull(Me.Date2) Then
        MsgBox "Please input the date", , "TB CMS"
        Date2.SetFocus
        Exit Sub
    End If
    'Me.Refresh
    'Form_frmClientLedger.Requery
    Form_frmClientLedger.Current_Balance.Requery
    'Me.Payment.SetFocus
    'Form_frmClientLedger.Requery
End Sub

Private Sub chkAdvancedLegal_Click()
If Me.AdvancedLegal = True Then
    Me.FirmPrepaid.value = False
End If


If Me.AdvancedLegal = False Then
    Me.FirmPrepaid.value = False
End If

Me.Refresh
End Sub


Private Sub chkAdvancedLegal_LostFocus()
    'Call reorderByDateMatter(Me.CaseID)
    'Debug.Print "aaaaaaaaaaaaaaaaaaaaaaa" & Form_frmClientLedger.CaseID
    Call reorderByDateMatter(Form_frmClientLedger.CaseID)
    Me.Requery
    Form_frmClientLedger.txtTrustBalance.Requery
    DoCmd.GoToRecord , , acNewRec
'    Dim strBookmark As String
'    strBookmark = Me.Form.Bookmark
'    Me.Form.Requery
'    Me.Form.Bookmark = strBookmark
End Sub

Private Sub cmdReceipt_Click()

'     If CurrentProject.AllReports("rptReceipt").IsLoaded Then
'        DoCmd.Close acReport, "rptReceipt", acSaveNo
'        End If
    If Me.Payment > 0 Then
    If CurrentProject.AllReports("rptReceiptC").IsLoaded Then
        DoCmd.Close acReport, "rptReceiptC", acSaveNo
    End If
        DoCmd.OpenReport "rptReceiptC", acViewPreview, , "[MatterID]=" & Me.MatterID
    End If
'    If Me.Pay_Outlay Like "*Cash*" Then
'    Reports!rptreceipt!OptionCash.value = True
'    End If
'    If Me.Pay_Outlay Like "*CC*" Then
'    Reports!rptreceipt!OptionCC.value = True
'    End If
'    If Me.Pay_Outlay Like "*Check*" Then
'    Reports!rptreceipt!OptionCheck.value = True
'    End If

End Sub

Private Sub FirmPrepaid_Click()
   
If Me.FirmPrepaid = True Then
    Me.AdvancedLegal.value = False
End If


If Me.FirmPrepaid = False Then
    Me.AdvancedLegal.value = False
End If

Me.Refresh
      
End Sub



Private Sub Form_AfterUpdate()
    Call reorderByDateMatter(Form_frmClientLedger.CaseID)
    Form_frmClientLedger.Current_Balance.Requery
End Sub

Private Sub Pay_Outlay_LostFocus()
    If Me.Pay_Outlay Like "Payment*" Then
    Payment.SetFocus
    End If
'    If Me.Pay_Outlay = "Payment (Cash)" Then
'    Payment.SetFocus
'    End If
'    If Me.Pay_Outlay Like "Payment*" Then
'    Payment.SetFocus
'    End If
End Sub

Private Sub Payment_AfterUpdate()
    'Me.Recalc
    If Me.Dirty Then Me.Dirty = False
    'Me.FirmPrepaid.SetFocus
    If IsNull(Me.Date2) Then
        MsgBox "Please input the date", , "TB CMS"
        Date2.SetFocus
        Exit Sub
    End If
    'Me.Refresh
    'Form_frmClientLedger.Requery
    Form_frmClientLedger.Current_Balance.Requery
    'Me.FirmPrepaid.SetFocus
    Me.txtTransPymtTrust.SetFocus
End Sub

'Private Sub Date2_Enter()
'    OldDate = Me.Date2
'    OldMatterID = Nz(Me.MatterID, 0)
'    OldPay_Outlay = Nz(Me.Pay_Outlay, 0)
'    OldCharge = Me.Charge
'    OldPayment = Me.Payment
'    OldFirmPrepaid = Me.FirmPrepaid
'End Sub

Private Sub Date2_AfterUpdate()
    
'    If OldDate <> Me.Date2 Then
'        'MsgBox "New date"
'        CurrentDb.Execute "delete * from [Matter and AR] where MatterID = " & OldMatterID
'        strInsert = "insert into [Matter and AR] (CaseID,Date2,Pay_Outlay,Charge,Payment,FirmPrepaid) " & _
'                    " values(" & Form_frmClientLedger.CaseID & ",#" & OldDate & "#,'" & OldPay_Outlay & "', " & OldCharge & "," & OldPayment & "," & OldFirmPrepaid & ")"
'        Debug.Print strInsert
'        CurrentDb.Execute strInsert
'    End If
    
    'Form_frmClientLedger.Requery
    'MsgBox dateOldVal
    'Form_frmClientLedger.Current_Balance.Requery
'    Me.Refresh
    'Dim lMatterID As Integer
    
    'lMatterID = Nz(Me.MatterID, 0)
    
    'DoCmd.RunCommand acCmdSaveRecord
    'Call reorderByDateMatter(Form_frmClientLedger.CaseID)
    
    'Me.Requery
    'Me.Refresh

    
    
    'Me.Recordset.FindFirst "MatterID = " & CStr(lMatterID)
    'DoCmd.GoToControl "Pay_Outlay"
    
'    Dim crId As Integer
'    crId = Me.CurrentRecord
'    Me.Requery
'    DoCmd.GoToRecord , , acGoTo, crId
    
'    Dim strBookmark As String
'    strBookmark = Me.Bookmark
'    Me.Requery
'    Me.Bookmark = strBookmark
    
    
    
    
    'Me.Requery
    
    'Me.Recalc
    'Form_frmClientLedger.Recordset.MoveNext
    'Form_frmClientLedger.Recordset.MovePrevious
    'Me.Refresh
End Sub



Private Sub txtTransPymtTrust_Click()
    If Me.InsertPymt = 0 Then
        Call cmdInsertIntoTA_Click
    End If
End Sub

Sub reorderByDateMatter(lngCaseID As Long)
    Dim rs As Recordset
    
    strSQL = "select * from [Matter and AR] where CaseID = " & lngCaseID & " order by Date2, Charge desc"
    Debug.Print strSQL
    
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
    Dim i As Integer
    i = 0
    Do Until rs.EOF
        i = i + 1
        'Take off table:
        strSQL = "Update [Matter and AR] set OrderNr = " & i & " where MatterID =" & rs("MatterID")
        Debug.Print "Date is: " & rs("date2")
        Debug.Print strSQL
        CurrentDb.Execute strSQL, dbSeeChanges
        
        rs.MoveNext
    Loop
    'MsgBox "Done", , "TB CMS"
End Sub

Private Sub cmdInsertIntoTA_Click()
    Dim strSQL As String
    
    If Me.Payment > 0 Then
        If Me.Pay_Outlay Like ("Payment (Cash*") And Me.Payment > 0 Then
            strSQL = "Insert into [Trust Account] (CaseID, Debit, TDate, TMatter)"
            strSQL = strSQL & " values (" & _
                              Me.txtCaseID & ", " & _
                              Nz(Me.Payment, 0) & ", #" & _
                              Date2 & "#, 'Deposit (Cash)')"
        End If
        
        If Me.Pay_Outlay Like ("Payment (Chk*") And Me.Payment > 0 Then
            strSQL = "Insert into [Trust Account] (CaseID, Debit, TDate, TMatter)"
            strSQL = strSQL & " values (" & _
                              Me.txtCaseID & ", " & _
                              Nz(Me.Payment, 0) & ", #" & _
                              Date2 & "#, 'Deposit (Chk)')"
        End If
                
        If Me.Pay_Outlay = "Payment (CC)" Then
            strSQL = "Insert into [Trust Account] (CaseID, Debit, TDate, TMatter)"
            strSQL = strSQL & " values (" & _
                              Me.txtCaseID & ", " & _
                              Nz(Me.Payment, 0) & ", #" & _
                              Date2 & "#, 'Deposit (CC)')"
        End If
    End If
        
    Debug.Print strSQL
    If Len(strSQL) > 10 Then
        CurrentDb.Execute strSQL
        i = 1
    End If
    
    Me.InsertPymt = -1
    'Me.Form.Recordset.Edit
    'Me.Form.Recordset("InsertPymt") = -1
    'Me.Form.Recordset.Update
    Me.Dirty = False
    
    Form_frmTrustAccount.reorderByDateTrustAccount
    
    Form_frmTrustAccount.Requery
    Form_frmClientLedger.txtTrustBalance.Requery
    Me.Requery
    Me.Refresh
    DoCmd.GoToRecord , , acNewRec
End Sub
