' Component: Form_frmTrustAccount
' Type: document
' Lines: 205
' ============================================================

Option Compare Database

Private Sub CheckCashed_AfterUpdate()
    If Me.Reconciled Then
        Me.Undo
    End If
End Sub

Private Sub Credit_AfterUpdate()
    'Me.Recalc
    'If Me.Dirty Then Me.Dirty = False
    'Form_frmClientLedger.txtTrustBalance.Requery
'    Dim lTrustAccountID As Integer
'    lTrustAccountID = Nz(Me.TrustAccountID, 0)
'
'
'    'If Me.Dirty Then Me.Dirty = False
'    DoCmd.RunCommand acCmdSaveRecord
'    Me.Requery
'    Me.Refresh
'    Me.Recordset.FindFirst "TrustAccountID = " & CStr(lTrustAccountID)
    
    'Me.CheckNumber.SetFocus
    'Form_frmClientLedger.Requery
    If Me.Dirty Then Me.Dirty = False
    Form_frmClientLedger.txtTrustBalance.Requery

End Sub

Private Sub Debit_AfterUpdate()
    'Me.Recalc
    'Me.Refresh
       
    'Dim lTrustAccountID As Integer
    'lTrustAccountID = Nz(Me.TrustAccountID, 0)
    
    
    'If Me.Dirty Then Me.Dirty = False
    'DoCmd.RunCommand acCmdSaveRecord
    'Me.Requery
    'Me.Refresh
    'Me.Recordset.FindFirst "TrustAccountID = " & CStr(lTrustAccountID)
    
    'Me.Credit.SetFocus
    'Form_frmClientLedger.Requery
    'DoCmd.GoToControl "Credit"
    
    If Me.Dirty Then Me.Dirty = False
    Form_frmClientLedger.txtTrustBalance.Requery

End Sub

Private Sub DepCleared_AfterUpdate()
    If Me.Reconciled Then
        Me.Undo
    End If
End Sub

Private Sub DepCleared_LostFocus()
    Call reorderByDateTrustAccount
    Me.Requery
    DoCmd.GoToRecord , , acNewRec
End Sub
Private Sub Form_AfterUpdate()
    Call reorderByDateTrustAccount
    Form_frmClientLedger.txtTrustBalance.Requery
End Sub

Private Sub Form_Current()
    'Added By RBD 10/6/17

   ReconciledEnableControl Me.TDate
   ReconciledEnableControl Me.CheckNumber
   ReconciledEnableControl Me.DepCleared
   ReconciledEnableControl Me.CheckCashed
   ReconciledEnableControl Me.Credit
   ReconciledEnableControl Me.Debit
    
End Sub
Public Sub ReconciledEnableControl(ctl As Access.Control)
    'Added By RBD 10/6/17
    'If Reconciled is True, lock and disable the passed control.  And vice versa.
    
    ctl.Locked = Nz(Me.Reconciled, False)
    ctl.Enabled = Not Nz(Me.Reconciled, False)
    
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler_cmdDelete_Click
    DoCmd.RunCommand acCmdDeleteRecord

cmdDelete_Click_Exit:
    Exit Sub
ErrHandler_cmdDelete_Click:
    If Err.Number <> 2501 Then  'Ignore any "action was cancelled" errors
       ShowMessage Err.Description
    End If
    Resume cmdDelete_Click_Exit


'RBD Note 10/7/17 - the below logic was changed based on PM's instructions, and moved to the BeforeDelConfirm event.

 '   answer = InputBox("Please input the admin password", "Admin pass required")
 '   If answer = "Paul123" And Me.Reconciled = True Then
 '       CurrentDb.Execute "delete * from [Trust Account] where TrustAccountID = " & Me.TrustAccountID
  '      Me.Requery
  '      DoCmd.GoToRecord , , acNewRec
 '   Else
 '       Beep
 '   End If
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    If Me.Reconciled = True Then
        answer = InputBox("Please input the admin password", "Admin pass required")
        If answer <> "Paul123" Then
            Cancel = True  'Cancel the deletion
            Beep
        End If
    End If
End Sub


'Private Sub Form_Current()
'    Me.TDate.Enabled = Not Me.Reconciled
'End Sub

' not working:
'Private Sub Form_Current()
'    If IsNull(Me.OrderNr) Then
'        Call reorderByDate
'        Me.Requery
'    End If
'End Sub

Private Sub TDate_AfterUpdate()
'    If Reconciled = -1 Then
'    MsgBox "Record has been reconciled.  Please contact administrator to make change.", , "TB CMS"
'    End Sub
'Else
'    DoCmd.RunCommand acCmdSaveRecord
'    Call reorderByDateTrustAccount
'    Me.Requery
'    Me.Refresh
'
'    DoCmd.GoToControl "TMatter"
    
    'Form_frmClientLedger.Current_Balance.Requery
    'If Me.Dirty Then Me.Dirty = False
    'Form_frmClientLedger.txtTrustBalance.Requery
    'Me.Requery
'    End If
End Sub

Sub reorderByDateTrustAccount()
    Dim rs As Recordset
    If Not IsNull(Form_frmClientLedger.CaseID) Then
        strSQL = "select * from [Trust Account] where CaseID = " & Form_frmClientLedger.CaseID & " order by TDate, Debit desc"
        'Debug.Print strSQL
    
        Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)

        Dim i As Integer
        i = 0
        Do Until rs.EOF
            i = i + 1
            strSQL = "Update [Trust Account] set OrderNr = " & i & " where TrustAccountID =" & rs("TrustAccountID")
            'Debug.Print "Date is: " & rs("tdate")
            'Debug.Print strSQL
            CurrentDb.Execute strSQL, dbSeeChanges
            rs.MoveNext
        Loop
    End If
End Sub

'Private Sub TMatter_AfterUpdate()
'    DoCmd.RunCommand acCmdSaveRecord
'    Me.Requery
'    Me.Refresh
'End Sub

'Private Sub cmdDelete_Click()
'    answer = InputBox("Please input the admin pasword", "Admin pass required")
'    If answer = "Paul123" And Me.Reconciled = True Then
'        CurrentDb.Execute "delete * from [Trust Account] where TrustAccountID = " & Me.TrustAccountID
'        Me.Requery
'        DoCmd.GoToRecord , , acNewRec
'    Else
'        Beep
'    End If
'End Sub

Private Sub TMatter_BeforeUpdate(Cancel As Integer)
    If Me.Reconciled = True Then
        answer = MsgBox("Are you sure you want to change the description?  This record has already been reconciled.", vbYesNo, "TB CMS")
        If answer = vbYes Then
            
        Else
            Cancel = True
            Me.Undo
        End If
    End If
End Sub
