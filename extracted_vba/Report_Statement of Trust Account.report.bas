' Component: Report_Statement of Trust Account
' Type: document
' Lines: 14
' ============================================================

Option Compare Database

Private Sub Report_Load()
    If StatementofTrustAccount_Filter = False Then
            Me.Filter = ""
            Me.FilterOn = False
    End If
End Sub

Private Sub Report_NoData(Cancel As Integer)
    Cancel = True
End Sub

