' Component: Form_frmPersInjProvider
' Type: document
' Lines: 17
' ============================================================

Option Compare Database

Private Sub cmdOpenDocumentFolderMedDocs_Click()
If Not OpenDocumentFolder([Forms]![frmClientLedger]![Me.CaseID], "Client Medical Records") Then
        MsgBox "Failed to open folder...", , "TB CMS"
    End If
End Sub

Private Sub cmdMedDocsfolder_Click()
    Dim lCaseID As Variant
    
    lCaseID = Forms!frmClientLedger!CaseID
    
    If Not OpenDocumentFolder(lCaseID, "Client Medical Records") Then
        MsgBox "Failed to open document folder...", , "TB CMS"
    End If
End Sub