' Component: Form_frmActionNeeded
' Type: document
' Lines: 68
' ============================================================

Option Compare Database

Private Sub cmdOutlookTask_Click()
    If Not IsNull(Me.ActionNeededDet) Then
  Const olFolderTasks = 13
    Dim olkApp, olkSes, olkFolder
    'On Error Resume Next
    'Set olkApp = GetObject(, "Outlook.Application")
    'If TypeName(olkApp) = "Nothing" Then
        Set olkApp = CreateObject("Outlook.Application")
        Set olkSes = olkApp.GetNameSpace("MAPI")
        'Change Outlook on the next line to the name of the default mail profile'
        
        
        ''''olkSes.Logon "Paul Mickelsen- Office 365"
        
        
    'Else
        'Set olkSes = olkApp.Session
    'End If
    Set olkFolder = olkSes.GetDefaultFolder(olFolderTasks)
    olkFolder.Display

    Dim objTask As Outlook.TaskItem
    Set objTask = olkApp.CreateItem(olTaskItem)

        With objTask
            .StartDate = Date
            .DueDate = Me.DateComp1
            '.Duration = 180 'Me!ApptLength
            .Subject = Form_frmClientLedger.Last_Name & ", " & Form_frmClientLedger.First_Name & " (" & Form_frmClientLedger.txtFileNo & "): " & Me.ActionNeededDet
            'LastName, FirstName (HearingType)
            '.Location = Form_frmClientLedger.Court

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
    Else
    MsgBox "Please enter Action Needed.", , "TB CMS"
    End If
End Sub