' Component: Form_frmPersInjLog
' Type: document
' Lines: 13
' ============================================================

Option Compare Database

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
 If Not Me.Dirty Then
   If (Count < 0) And (Me.CurrentRecord > 1) Then
     DoCmd.GoToRecord , , acPrevious
   ElseIf (Count > 0) And (Me.CurrentRecord < Me.Recordset.RecordCount) Then
        DoCmd.GoToRecord , , acNext
   End If
 Else
   MsgBox "The record has changed. Save the current record before moving to another record.", , "TB CMS"
 End If
End Sub