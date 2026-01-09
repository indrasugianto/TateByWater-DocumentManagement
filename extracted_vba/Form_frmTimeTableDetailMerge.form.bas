' Component: Form_frmTimeTableDetailMerge
' Type: document
' Lines: 95
' ============================================================

Option Compare Database

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    Response = 0
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "TB CMS") = vbNo Then Cancel = True
End Sub

'Private Sub Tdate_BeforeUpdate(Cancel As Integer)
'    If Not IsNull(Me.TDate) Then
'        If Me.TDate > Me.Parent.Bill_Sent Then
'            Mbox "Date", 2
'            Cancel = True
'        End If
'    Else
'        Mbox "Date", 1
'        Cancel = True
'    End If
'End Sub
'
'Private Sub Time__AfterUpdate()
'    Me.Recalc
'    Me.Rate.SetFocus
'End Sub
'Private Sub Rate_AfterUpdate()
'    Me.Recalc
'    Me.Tatty.SetFocus
'End Sub

'Private Sub Form_AfterDelConfirm(Status As Integer)
'    Call RefreshForms
'End Sub
'
'Private Sub Form_AfterInsert()
'    Call RefreshForms
'End Sub
'
'Private Sub Form_AfterUpdate()
'    MsgBox "Refresh forms AAAA"
'    Call RefreshForms
'End Sub
'
'Private Sub Form_Delete(Cancel As Integer)
'    Call RefreshForms
'End Sub

'Sub RefreshForms()
'    'MsgBox "Refresh forms"
'
'    'Me.Recalc
'    '[Form_Time Keeping].Recalc
'
'    '[Forms]![Form_Time Keeping].Refresh
'
'    '[Form_Time Keeping].Requery
'    'Forms![frmProject].[Form].[txtTotalTime].Requery
'
'
'    Me.Requery
'    [Form_Time Keeping].Requery
'
'End Sub
'
'Private Sub TxtRunningTotal_AfterUpdate()
'    Call RefreshForms
'End Sub
'
'Private Sub TxtRunningTotal_BeforeUpdate(Cancel As Integer)
'    Call RefreshForms
'End Sub
'
'Private Sub TxtRunningTotal_Change()
'    Call RefreshForms
'End Sub
'
'Private Sub Rate_AfterUpdate()
'    Call RefreshForms
'End Sub

Private Sub Form_Timer()
    Me![txttenths].Visible = False      'make the Text Box invisible after 5 secs
    Me.TimerInterval = 0              'disable the Timer
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    If IsNull([Form_Time Keeping].CaseID) Then
        MsgBox "Please select an existing TK or add a new TK", , "TB CMS"
        [Form_Time Keeping].cmbBills.SetFocus
        Cancel = True
    End If
End Sub

Private Sub cmdTenths_Click()
    Me![txttenths].Visible = True     'make Text Box invisible
    Me.TimerInterval = 10000         'set Timer Interval to 10 seconds
End Sub