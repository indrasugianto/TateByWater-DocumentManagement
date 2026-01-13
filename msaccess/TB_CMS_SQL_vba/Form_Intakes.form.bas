' Component: Form_Intakes
' Type: document
' Lines: 287
' ============================================================

Option Compare Database
'Private cls As New clsFormValidation
'Public Err3314Encountered As Boolean


Private Sub cmdClearFilter_Click()
    Me.FilterOn = False
End Sub

Private Sub cmbHome_Click()
DoCmd.openform "frmhome", acNormal
End Sub

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub



'Private Sub cmdCopyToClient_Click()
'On Error GoTo ErrHnd
'Dim Var As Variant
'Dim rs As Recordset
'Dim intval As Long
'
'
'            If Nz(Me.ID, 0) <> 0 Then
'                If DCount("ID", "TB Intakes", "ID=" & Me.ID) = 0 Then
'                    If MsgBox("The record is not saved. Would you like the application to save it automatically? (If not then please save it first)", vbYesNo + vbQuestion, "Unsaved Record") = vbYes Then
'                        cls.ExeCommand SaveRec
'                    Else
'                        Exit Sub
'                    End If
'                Else
'                ' Do nothing and continue
'                End If
'                Set rs = CurrentDb.OpenRecordset("Select * from [TB Client Info]")
'                    rs.AddNew
'                    With rs
'                        !Last_Name = Nz(Me.GI_Last_Name, "")
'                        !First_Name = Nz(Me.GI_First_Name, "")
'                        !HmPhone = Nz(Me.GI_Phone, "")
'                        !Referral = Nz(Me.Referral, "")
'                        ![Individual Referrer] = Nz(Me.GI_Individual_Referrer, "")
'
'                    End With
'                    rs.Update
'                    Var = rs.LastModified
'                    rs.Bookmark = Var
'                CurrentDb.Execute "Insert into tblCase (CaseID,Case_Letter,CaseOpenDate) Values (" & Nz(rs!CaseID, 0) & ",'" & Me.GI_Practice_Area & "','" & Me.Date_Opened & "')", dbFailOnError
'                CaseGenerator '==== to generate case number year wise
'                CurrentDb.Execute ("DELETE * FROM [TB Intakes] WHERE ID =" & Me.ID) '=== Deleting the record from tbl intakes
'                Me.Requery
'                MsgBox "Data successfully copied to Client Info table.", vbInformation
'                 rs.Close
'                 Set rs = Nothing
'            Else
'                MsgBox "No record found!", vbInformation, "Invalid Command"
'            End If
'
'
'Exit Sub
'ErrHnd:
'    ErrMsg "cmdCopyToClient_Click"
'End Sub

Private Sub Command227_Click()
    'cls.ExeCommand Cancelrec
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close
End Sub


'Private Sub Form_Activate()
'    If Err3314Encountered = True Then
'       Me.Controls(strCtlNameRaw).BackColor = RGB(255, 0, 0)
'    End If
'End Sub



'Private Sub Form_Error(DataErr As Integer, Response As Integer)
'    '==== CODE TO HANDLE THE ACCESS ERROR FOR BLANK MANDATORY FIELD.
'        Dim strCtlName          As String
'        Dim Ctl                 As Control
'        For Each Ctl In Application.Forms("Intakes")
'            If Ctl.Tag = "Req" Then
'                If IsNull(Ctl) Then
'
'                    strCtlName = Right(Ctl.name, Len(Ctl.name) - InStr(1, Ctl.name, " ", vbBinaryCompare))
'                    Err3314Encountered = True
'                    GoTo BlankControlFound
'                End If
'            ElseIf Ctl.Tag = "Req_" Then
'                If IsNull(Ctl) Then
'
'                    strCtlName = Ctl.name
'                    Err3314Encountered = True
'                    GoTo BlankControlFound
'                End If
'            End If
'        Next
'
'BlankControlFound:
'            Select Case DataErr
'                Case 3314:
'                            MsgBox "The field named: " & strCtlName & " cannot be left blank.", vbCritical, "Mandatory Field Missing"
'                            Response = acDataErrContinue
'            End Select
'End Sub

Private Sub Referral_AfterUpdate()
'    If Nz(Me.Referral.Column(1), "") = "Individual Referral" Then
'        Me.GI_Individual_Referrer.Enabled = True
'        Me.GI_Individual_Referrer.SetFocus
'    Else
'        Me.GI_Individual_Referrer.Enabled = False
'    End If
End Sub

Private Sub Command225_Click()
'    If Me.Dirty Then
'        cls.ExeCommand SaveRec
'    End If
    'DoCmd.RunCommand acCmdSaveRecord
    If Me.Dirty Then Me.Dirty = False
End Sub


Private Sub cmdCreateOpen_Click()
'    On Error GoTo ErrHnd
'    Dim Var As Variant
'    Dim intval As Long
'
'           If Nz(Me.ID, 0) <> 0 Then
'                If DCount("ID", "TB Intakes", "ID=" & Me.ID) = 0 Then
'                    If MsgBox("The record is not saved. Would you like the application to save it automatically? (If not then please save it first)", vbYesNo + vbQuestion, "Unsaved Record") = vbYes Then
'                        cls.ExeCommand SaveRec
'                    Else
'                        Exit Sub
'                    End If
'                Else
'                ' Do nothing and continue
'                End If
'                Set rs = CurrentDb.OpenRecordset("Select * from tblcase")
'                    rs.AddNew
'                    With rs
'                        !Last_Name = Nz(Me.GI_Last_Name, "")
'                        !First_Name = Nz(Me.GI_First_Name, "")
'                        !HmPhone = Nz(Me.GI_Phone, "")
'                        !Referral = Nz(Me.GI_Referral, "")
'                        ![Individual Referrer] = Nz(Me.GI_Individual_Referrer, "")

'                    End With
'                    rs.Update
'                   Var = rs.LastModified
'                    rs.Bookmark = Var
'                Me.[GI Open Date] = Date
'                Me.[GI Open] = -1
'                CurrentDb.Execute "Insert into tblCase (CaseID,Case_Letter,CaseOpenDate, Last_Name, First_Name, HmPhone, Referral, Individual_Referrer) Values (" & Nz(rs!CaseID, 0) & ",'" & Me.GI_Practice_Area & "','" & Me.GI_Open_Date & "')", dbFailOnError
'                CaseGenerator '==== to generate case number year wise
'                'CurrentDb.Execute ("DELETE * FROM [TB Intakes] WHERE ID =" & Me.ID) '=== Deleting the record from tbl intakes
'                Me.Requery
'               MsgBox "Client File Open.", vbInformation
'                rs.Close
'                 Set rs = Nothing
'            Else
'               MsgBox "No record found!", vbInformation, "Invalid Command"
'            End If
'
'
'    Exit Sub
'ErrHnd:
'    ErrMsg "cmdCreateOpen_Click_Click"
End Sub

Private Sub cmdOpenIntakeDocument_Click()
    'check if the document exists
    If Not pcaempty([Scan Location GI]) Then
        If Not pcaempty(Dir([Scan Location GI])) Then
            Application.FollowHyperlink Scan_Location_GI
        End If
    End If
End Sub

Private Sub cmdScan_Click()
Dim ScannerDirectory As String
Dim SelectedFileName As String
Dim DestinationFileName As String
Dim FolderName As String

    If pcaempty(Me.ID) Then
        MsgBox "Please select the Intake record before continuing...", , "TB CMS"
    Else

        'select the scanned file
        ScannerDirectory = GetScannerFolder()
        SelectedFileName = SelectFileDialog("Select scanned file", ScannerDirectory, "")
        
        If Not pcaempty(SelectedFileName) Then
            'if the document type is Closed Final, need to put copy on the CLOSED FINAL SCAN folder as well
            
            DestinationFileName = GetIntakeDocumentFileName(ID)

            'append file extension to the destination file
            DestinationFileName = DestinationFileName & "." & Right(SelectedFileName, Len(SelectedFileName) - InStrRev(SelectedFileName, "."))
            
            FolderName = GetIntakeFolderName()

            'check if the folder exists, if not create one
            If FolderExistsCreate(FolderName, True) Then
                Set fDialog = Application.FileDialog(msoFileDialogSaveAs)
        
            With fDialog
                .AllowMultiSelect = False
                .InitialFileName = FolderName
                
                .InitialFileName = FolderName & DestinationFileName
                
                
                If .show = True Then
                    For Each varFile In .SelectedItems
                        'save the file
                        FileCopy SelectedFileName, varFile
                        
                        'store the file location on Intake table
                        Me.Scan_Location_GI = varFile
                        Me.Scanned = True
                        Me.Dirty = False
                        MsgBox "Intake document sucessfully scanned", , "TB CMS"
                    Next
                End If
            End With
        End If

        End If
    End If
End Sub

'Private Sub Form_Load()
'    'Set cls.Form = Me
'    'DoCmd.RunCommand acCmdRecordsGoToNew
'    'cls.RequiredControls = Array(Me.GI_Last_Name, Me.GI_First_Name, Me.GI_Phone, Me.Referral, Me.GI_Practice_Area, Me.Date_Opened)
'    'cls.RequiredControls = Array(Me.GI_Last_Name, Me.GI_First_Name, Me.Date_Opened)
'End Sub

Private Sub Command226_Click()
    'cls.ExeCommand Addrec
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub cmdLastWeekIntakes_Click()
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim lngDayNum As Long

    If Me.Dirty Then   '==== to check if some unsaved changes are there
        If MsgBox("Would  you like to save the changes?", vbYesNo + vbQuestion, "TB CMS: Save Changes") = vbYes Then
            DoCmd.RunCommand acCmdSaveRecord
        Else
            Me.Undo
        End If
    End If

    lngDayNum = Weekday(Date)
    dtEnd = DateAdd("d", -lngDayNum, Date)
    dtStart = DateAdd("d", -6, dtEnd)

    dtST1 = dtStart
    dtST2 = dtEnd

    DoCmd.OpenReport "rptLastWeekIntake", acViewPreview
End Sub


'Private Sub Form_Load()
'    Me.filter = "[Gi Date] between #" & DateAdd("m", -6, Date) & "# And #" & Date & "#"
'    Me.FilterOn = True
'End Sub

Private Sub cmdFilterLast6months_Click()
    'DoCmd.ApplyFilter , "GiDate" = DateAdd("m", -6, Date) And Date
    Me.Filter = "[Gi Date] between #" & DateAdd("m", -6, Date) & "# And #" & Date & "#"
    Me.FilterOn = True
End Sub

