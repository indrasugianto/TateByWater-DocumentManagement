' Component: Form_frmHearingDate
' Type: document
' Lines: 158
' ============================================================

 Option Compare Database


'Private Sub cmdAddAppt_Click()
'    On Error GoTo Add_Err
'
'    'Save record first to be sure required fields are filled.
'    DoCmd.RunCommand acCmdSaveRecord
'
'    'Exit the procedure if appointment has been added to Outlook.
'    If Me!AddedToOutlook = True Then
'        MsgBox "This appointment is already added to Microsoft Outlook"
'        Exit Sub
'    'Add a new appointment.
'    Else
'        Dim objOutlook As Outlook.Application
'        Dim objAppt As Outlook.AppointmentItem
'        Dim objRecurPattern As Outlook.RecurrencePattern
'
'        Set objOutlook = CreateObject("Outlook.Application")
'        Set objAppt = objOutlook.CreateItem(olAppointmentItem)
'
'        With objAppt
'            .Start = Me!Hearing_Date & " " & Me!HearingTime
''            .Duration = Me!ApptLength
'            .Subject = Me!HearingType
'            .Location = Me!Court
'
'            If Not IsNull(Me!ApptNotes) Then .Body = Me!ApptNotes
'            If Not IsNull(Me!ApptLocation) Then .Location = Me!ApptLocation
'            If Me!ApptReminder Then
'                .ReminderMinutesBeforeStart = Me!ReminderMinutes
'                .ReminderSet = True
'            End If
'
'            Set objRecurPattern = .GetRecurrencePattern
'
'            With objRecurPattern
'                .RecurrenceType = olRecursWeekly
'                .Interval = 1
'                'Once per week
'                .PatternStartDate = #7/9/2003#
'                'You could get these values
'                'from new text boxes on the form.
'                .PatternEndDate = #7/23/2003#
'            End With
'
'            .Save
'            .Close (olSave)
'            End With
'            'Release the AppointmentItem object variable.
'            Set objAppt = Nothing
'    End If
'
'    'Release the Outlook object variable.
'    Set objOutlook = Nothing
'
'    'Set the AddedToOutlook flag, save the record, display a message.
'    Me!AddedToOutlook = True
'    DoCmd.RunCommand acCmdSaveRecord
'    MsgBox "Appointment Added!"
'
'    Exit Sub
'
'Add_Err:
'    MsgBox "Error " & Err.Number & vbCrLf & Err.Description
'    Exit Sub
'End Sub
'
Private Sub cmdOutlookCalendar2_Click()
    
    If IsNull(Me.Hearing_Date) Then
        MsgBox "Please input a correct date", , "TB CMS"
        Me.Hearing_Date.SetFocus
        Exit Sub
    End If
    
    Const olFolderCalendar = 9
    Dim olkApp, olkSes, olkFolder
    'On Error Resume Next
    'Set olkApp = GetObject(, "Outlook.Application")
    'If TypeName(olkApp) = "Nothing" Then
        Set olkApp = CreateObject("Outlook.Application")
        Set olkSes = olkApp.GetNameSpace("MAPI")
        'Change Outlook on the next line to the name of the default mail profile'
        'olkSes.Logon "Paul Mickelsen- Office 365"
    'Else
        'Set olkSes = olkApp.Session
    'End If
    Set olkFolder = olkSes.GetDefaultFolder(olFolderCalendar)
    olkFolder.Display

    Dim objAppt As Outlook.AppointmentItem
    Set objAppt = olkApp.CreateItem(olAppointmentItem)

        With objAppt
            .Start = Me.Hearing_Date & " " & Me.HearingTime
            .Duration = 120 'Me!ApptLength
            
            strHearingType = ""
            If Not IsNull(Me.HearingType) Then
                strHearingType = " (" & Me.HearingType & ")"
            End If
            
            .Subject = Form_frmClientLedger.Last_Name & ", " & Form_frmClientLedger.First_Name & strHearingType
            'LastName, FirstName (HearingType)
            '.Location = Nz(Form_frmClientLedger.CourtandType, "") '& " " & Form_frmClientLedger.CType
            strCourtAndType = Nz(Form_frmClientLedger.cmbCourt, "") & " " & Nz(Form_frmClientLedger.cmbCType, "")
            ' Nz(Me.cmbCourt, 0) + " " + Nz([CType], 0)
            .Location = strCourtAndType '& " " & Form_frmClientLedger.CType
            
'            If Not IsNull(Me!ApptNotes) Then .Body = Me!ApptNotes
'            If Not IsNull(Me!ApptLocation) Then .Location = Me!ApptLocation
'            If Me!ApptReminder Then
'                .ReminderMinutesBeforeStart = Me!ReminderMinutes
'                .ReminderSet = True
'            End If
'
'            Set objRecurPattern = .GetRecurrencePattern
'
'            With objRecurPattern
'                .RecurrenceType = olRecursWeekly
'                .Interval = 1
'                'Once per week
'                .PatternStartDate = #7/9/2003#
'                'You could get these values
'                'from new text boxes on the form.
'                .PatternEndDate = #7/23/2003#
'            End With
            .Display
 '           .Save
            '.Close (olSave)
        End With
            'Release the AppointmentItem object variable.
            'Set objAppt = Nothing


    'close outlook
    Set olkApp = Nothing
    Set olkSes = Nothing
    Set olkFolder = Nothing
    
    Me.HrgCal.value = -1
    
End Sub

Private Sub Combo40_AfterUpdate()
    If ClientPresent = 0 Then
        Me.Reminder.value = Null
    End If
End Sub

Private Sub Combo40_BeforeUpdate(Cancel As Integer)
    If ClientPresent = 0 Then
        Me.Reminder.value = Null
        MsgBox "No need to send reminder of hearing where client not required to be present.", vbInformation, "TB CMS"
    End If
End Sub