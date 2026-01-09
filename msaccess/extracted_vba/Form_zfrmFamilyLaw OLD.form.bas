' Component: Form_zfrmFamilyLaw OLD
' Type: document
' Lines: 92
' ============================================================

Option Compare Database

Private cls As New clsFormValidation

Private Sub cmdClose_Click()
    'cls.ExeCommand Cancelrec
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Command405_Click()
    cls.ExeCommand SaveRec
End Sub

Private Sub Command406_Click()
    cls.ExeCommand Addrec
End Sub


Private Sub D_DOB_BeforeUpdate(Cancel As Integer)
    If Not IsNull(Me.D_DOB) Then
        Cancel = DateVarifier(Me.D_DOB)
    Else
        Mbox "D_DOB", 1
        Cancel = True
    End If
End Sub

Private Sub Date_of_Marriage_BeforeUpdate(Cancel As Integer)
    If Not IsNull(Me.Date_of_Marriage) Then
        Cancel = DateVarifier(Me.Date_of_Marriage)
    Else
        Mbox "Date of marriage", 1
        Cancel = True
    End If
End Sub

Private Sub Date_of_Separation_BeforeUpdate(Cancel As Integer)
    If Not IsNull(Me.Date_of_Separation) Then
        If IsNull(Me.Date_of_Marriage) Then
            MsgBox "Kindly enter Marriage date first.", vbInformation, "Invalid Input"
            Cancel = True
        Else
            If Me.Date_of_Separation < Me.Date_of_Marriage Then
                MsgBox "Date of separation can't be earlier than date of marriage.", vbInformation, "Invalid Input"
                Cancel = True
            Else
                Cancel = DateVarifier(Me.Date_of_Separation)
            End If
        End If
    Else
        Mbox "Date of separation", 1
        Cancel = True
    End If
End Sub

'Private Sub DOB_BeforeUpdate(Cancel As Integer)
'    If Not IsNull(DOB) Then
'        Cancel = DateVarifier(Me.DOB)
'    Else
'        Mbox "DOB", 1
'    Cancel = True
'    End If
'End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    Response = 0
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "TB CMS") = vbNo Then Cancel = True
End Sub

Private Sub Form_Load()
'    Set cls.Form = Me
'    cls.RequiredControls = Array(Me.Date_of_Marriage, Me.Date_of_Separation, Me.D_Home_Phone)
'    If lngChnc = 2 Then
'        Me.AllowAdditions = False
'    End If
'    lngChnc = 1
'    If DCount("ID", "Family Law - Divorce", "CaseID=" & Nz(Me.OpenArgs, 0)) = 0 Then
'        If Nz(Me.OpenArgs, "") <> "" Then
'            Me.CaseID = Nz(Me.OpenArgs, 0)
'        End If
'    Else
'        Me.FilterOn = False
'        Me.filter = "CaseID=" & Me.OpenArgs
'        Me.FilterOn = True
'    End If
End Sub

Private Sub cmbClients_AfterUpdate()
    Me.Filter = "[Family Law - Divorce].CaseID = " & Me.cmbClients
    Debug.Print Me.Filter
    Me.FilterOn = True
End Sub