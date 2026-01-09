' Component: clsFormValidation
' Type: class
' Lines: 95
' ============================================================

Option Compare Database

Public RequiredControls         As Variant    ' Array (me.Title,me.Author)
Private WithEvents frm          As Form

Public Enum ControlType
    SaveRec
    Addrec
    Cancelrec
End Enum

Private szCtlNameBackColor As String
Private lngCmdType As Long

Public Property Set Form(frmObject As Form)
    Set frm = frmObject
    If frm.BeforeUpdate = "" Then frm.BeforeUpdate = "[Event Procedure]"
End Property

Private Sub Frm_BeforeUpdate(Cancel As Integer)
    
    If Not IsEmpty(RequiredControls) Then
         If IsMandatoryByName(RequiredControls) Then
            Cancel = True
            GoTo endit
         End If
        For i = LBound(RequiredControls) To UBound(RequiredControls)
            frm.Controls(RequiredControls(i).Name).BackColor = vbWhite
        Next
    'If lngCmdType = 0 Then MsgBox "Record Saved.", vbInformation, "Save"
    End If
endit:
End Sub

Public Sub ExeCommand(Cmd As ControlType)
    Dim frm As Form

    Set frm = CodeContextObject
    On Error GoTo ErrHand
    If Cmd <> Cancelrec Then DoCmd.RunCommand acCmdSaveRecord
        
    If Cmd = SaveRec Then
        MsgBox "Record saved", vbInformation, "Save"
    ElseIf Cmd = Addrec Then
        If frm.AllowAdditions Then
           DoCmd.GoToRecord , , acNewRec
        Else
            MsgBox "You are not allowed to add new records to this form.", vbExclamation
        End If
    Else
        If frm.Dirty Then
            Select Case MsgBox("Do you want to save the changes?", vbQuestion + vbYesNoCancel, "Save Changes?")
                Case vbYes
              ''Great! Run the save command and let them know if there are any validation before they save.
                     DoCmd.RunCommand acCmdSaveRecord
                     'DoCmd.Close acForm, frm.Name, acSaveYes
                Case vbNo
                     frm.Undo
                     'DoCmd.Close acForm, frm.Name, acSaveNo
            End Select
        Else
            'DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    End If
        Set frm = Nothing
    Exit Sub
ErrHand:
    If Err.Number <> 2501 And Err.Number <> 3021 Then MsgBox Err.Description
End Sub

Private Function IsMandatoryByName(ByVal VarControls As Variant) As Boolean
    Dim ctl         As Control
    Dim MyArray()   As Variant
    Dim MSG         As String
    Dim Counter     As Long
    
    For i = LBound(VarControls) To UBound(VarControls)
        
        Select Case VarControls(i).ControlType
            Case 109, 111
                Set ctl = frm.Controls(VarControls(i).Name)
            Case Else
                MsgBox "Control not recognized for this functionality.", vbExclamation
                Exit Function
        End Select
        
        If Nz(ctl, "") = "" Then
            ctl.BackColor = 13434879
            MsgBox "Field '" & ctl.Tag & "' can't be blank."
            IsMandatoryByName = True
            Exit Function
        End If
    Next
    
End Function