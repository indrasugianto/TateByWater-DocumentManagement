' Component: Form_frmCalls
' Type: document
' Lines: 392
' ============================================================

Option Compare Database
'Private cls As New clsFormValidation
'Public Err3314Encountered As Boolean

Private Sub CAttorney_AfterUpdate()
     If Me.cmbEmailAtty = "JRT" Then
        Me.attyEmail.value = "jtate@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "DEB" Then
        Me.attyEmail.value = "dbywater@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "GBF" Then
        Me.attyEmail.value = "gfuller@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "PM" Then
        Me.attyEmail.value = "pmickelsen@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "TDT" Then
        Me.attyEmail.value = "ttull@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "JN" Then
        Me.attyEmail.value = "jnolasco@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "BA" Then
        Me.attyEmail.value = "bader@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "KDB" Then
        Me.attyEmail.value = "kbigus@tatebywater.com"
    End If
     If Me.cmbEmailAtty = "WNE" Then
        Me.attyEmail.value = "wevans@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "RRL" Then
        Me.attyEmail.value = "rlucero@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "DK" Then
        Me.attyEmail.value = "dkrisky@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "RYR" Then
        Me.attyEmail.value = "rranaf@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "KP" Then
        Me.attyEmail.value = "kpizarro@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "RLF" Then
        Me.attyEmail.value = "rfredericks@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "CMH" Then
        Me.attyEmail.value = "cmhiggins@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "HG" Then
        Me.attyEmail.value = "hgrossman@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "NH" Then
        Me.attyEmail.value = "nhasan@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "AO" Then
        Me.attyEmail.value = "aorellana@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "IQ" Then
        Me.attyEmail.value = "iquintero@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "LV" Then
        Me.attyEmail.value = "lvelasquez@tatebywater.com"
    End If
    If Me.cmbEmailAtty = "JI" Then
        Me.attyEmail.value = "jilarraza@tatebywater.com"
    End If
    
    
End Sub
Private Sub cmbHome_Click()
DoCmd.openform "frmhome", acNormal
End Sub
Private Sub cmdBankruptcySend_Click()

    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Bankruptcy: Please call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Me.CallComments
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "kvasquez@tatebywater.com" & "; " & "gfuller@tatebywater.com"
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub

Private Sub cmdCallListForm_Click()
    DoCmd.openform "frmCallsList", acNormal
End Sub

Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdEstatePlanning_Click()

    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Estate Planning: Please call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "bader@tatebywater.com" & "; " & "kbigus@tatebywater.com"
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub

Private Sub cmdOutlookCalendar2_Click()
    Me.Refresh
    If IsNull(Me.SchedDate) Then
        MsgBox "Please input a correct date", , "TB CMS"
        Me.SchedDate.SetFocus
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

            strMatter = ""
            If Not IsNull(Me.CallMatter) Then
                strMatter = " (" & Me.CallMatter & ")"
            End If
            
            strSource = ""
            If Not IsNull(Me.CReferral) Then
                strSource = " (" & Me.CReferral & ")"
            End If
            
            strClientTYpe = ""
            If Not IsNull(Me.ClientType) Then
                strClientTYpe = ", " & Me.ClientType
            End If
            
            strPhone = ""
            If Not IsNull(Me.CPhone) Then
                strPhone = ", " & Me.CPhone
            End If
            
            strCPhoneType = ""
            If Not IsNull(Me.CPhoneType) Then
                strCPhoneType = " (" & Me.CPhoneType & ")"
            End If
            
        With objAppt
            .Start = Me.SchedDate & " " & Me.SchedTime
            .Duration = 60 'Me!ApptLength
            .Subject = "C: " & Me.CFirstName & " " & Me.CLastName & strPhone & strCPhoneType & strMatter & strSource & strClientTYpe
            .Body = "Invite sent: " & Nz(Me.CDate) & ",  " & Nz(Me.CallTime) & ",  " & Nz(Me.CallComments)
            .Display
        End With
            'Release the AppointmentItem object variable.
            'Set objAppt = Nothing
    'close outlook
    Set olkApp = Nothing
    Set olkSes = Nothing
    Set olkFolder = Nothing
End Sub

Private Sub cmdParalegals_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Paralegals: Please advise and call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "hgrossman@tatebywater.com" & "; " & "iquintero@tatebywater.com" & "; " & "aorellana@tatebywater.com" & "; " & "kpizarro@tatebywater.com" & "; " & "lvelasquez@tatebywater.com" & "; " & "jilarraza@tatebywater.com"
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub

Private Sub cmdPartners_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Partners. Please advise and/or call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "jtate@tatebywater.com" & "; " & "dbywater@tatebywater.com" & "; " & "cmhiggins@tatebywater.com" & "; " & "pmickelsen@tatebywater.com" & "; " & "ttull@tatebywater.com" & "; " & "nhasan@tatebywater.com" & "; " & "rfredericks@tatebywater.com" & "; " & "bader@tatebywater.com"
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub

Private Sub cmdSendAll_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Please call and/or advise: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "jtate@tatebywater.com" & "; " & "dbywater@tatebywater.com" & "; " & "wevans@tatebywater.com" & "; " & "pmickelsen@tatebywater.com" & "; " & "ttull@tatebywater.com" & "; " & "nhasan@tatebywater.com" & "; " & "bader@tatebywater.com" & "; " & "kbigus@tatebywater.com" & "; " & "cmhiggins@tatebywater.com" & "; " & "rfredericks@tatebywater.com" & "; " & "dkrisky@tatebywater.com" & "; " & "rlucero@tatebywater.com" & "; " & "rrana@tatebywater.com"
        .Subject = subject_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub

Private Sub cmdSendEmail_Click()

    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    email_ = Me.attyEmail
    subject_ = "Please call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter) & ", " & Nz(Me.ClientType)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = email_
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub
Private Sub cmdSendFamily_Click()
    
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Potential Family Law. Please advise and call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "jtate@tatebywater.com" & "; " & "dbywater@tatebywater.com" & "; " & "pmickelsen@tatebywater.com" & "; " & "ttull@tatebywater.com" & "; " & "nhasan@tatebywater.com" & "; " & "cmhiggins@tatebywater.com" & "; " & "rfredericks@tatebywater.com" & "; " & "hgrossman@tatebywater.com" & "; " & "kpizarro@tatebywater.com"
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
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
Private Sub Form_Load()
    DoCmd.GoToRecord , , acNewRec
End Sub

Private Sub SendCrim_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "Potential Criminal. Please advise and call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "jtate@tatebywater.com" & "; " & "pmickelsen@tatebywater.com" & "; " & "ttull@tatebywater.com" & "; " & "nhasan@tatebywater.com" & "; " & "wevans@tatebywater.com" & "; " & "aorellana@tatebywater.com"
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub

Private Sub cmdAllTBSend_Click()
    Me.Refresh
    Dim OutlookApp As Object
    Dim MItem As Object
    Dim email_ As String
    Dim subject_ As String
    Dim body_ As String
    Dim attach_ As String
     
    Set OutlookApp = CreateObject("Outlook.Application")
     
    subject_ = "ALL TB: Please advise and/or call: " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & ", " & Nz(Me.CallMatter)
    body_ = Me.CDate & ",  " & Me.CallTime & ": " & Nz(Me.CFirstName) & " " & Nz(Me.CLastName) & ", " & Nz(Me.CPhone) & Nz(Me.CallComments)
     
     'create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(0)
    With MItem
        .To = "jtate@tatebywater.com" & "; " & "dbywater@tatebywater.com" & "; " & "hgrossman@tatebywater.com" & "; " & "pmickelsen@tatebywater.com" & "; " & "ttull@tatebywater.com" & "; " & "nhasan@tatebywater.com" & "; " & "bader@tatebywater.com" & "; " & "kbigus@tatebywater.com" & "; " & "rfredericks@tatebywater.com" & "; " & "aorellana@tatebywater.com" & "; " & "iquintero@tatebywater.com" & "; " & "kpizarro@tatebywater.com" & "; " & "kmcallister@tatebywater.com" & "; " & "wevans@tatebywater.com" & "; " & "cmhiggins@tatebywater.com" & "; " & "rlucero@tatebywater.com" & "; " & "jnolasco@tatebywater.com" & "; " & "rrana@tatebywater.com" & "; " & "lvelasquez@tatebywater.com" & "; " & "jilarraza@tatebywater.com"
        .Subject = subject_
        .Body = body_
         '.Attachments.Add "C:\FolderName\Filename.txt"
        .Display
    End With
End Sub