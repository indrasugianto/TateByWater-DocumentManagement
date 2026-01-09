' Component: Form_frmChild
' Type: document
' Lines: 16
' ============================================================

Option Compare Database

Private Sub DOB_child_BeforeUpdate(Cancel As Integer)
    If Not IsNull(Me.DOB_child) Then
        Cancel = DateVarifier(Me.DOB_child)
    Else
        Mbox "DOB Child", 1
        Cancel = True
    End If
    
End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    Response = 0
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "TB CMS") = vbNo Then Cancel = True
End Sub