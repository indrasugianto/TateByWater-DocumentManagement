' Component: OutlookApp
' Type: class
' Lines: 170
' ============================================================

Option Compare Database
Private olApp As Object
Private pInboxPath As String
Private pAttachments As Collection



Public Property Get inboxpath() As String
    inboxpath = pInboxPath
End Property

Public Property Get attachments() As Collection
    Set attachments = pAttachments
End Property

Public Property Let inboxpath(value As String)
    pInboxPath = value
End Property


Public Sub CreateApp()
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
End Sub

Public Sub AddAttachments(strAttachment As String)
    If pAttachments Is Nothing Then
        Set pAttachments = New Collection
    End If
    pAttachments.Add strAttachment
End Sub

Public Sub SendMail(strMailBody As String, strSubject As String, strTo As String, Optional strCc As String = "", Optional strBcc As String = "")
    Dim olMail As Object
    Dim itm
    Set olMail = olApp.CreateItem(0)
    
    With olMail
        .Body = strBody
        .Subject = strSubject
        .To = strTo
        If strCc <> "" Then
            .CC = strCc
        End If
        
        If strBcc <> "" Then
            .BCC = strBcc
        End If
        If Me.attachments.Count > 0 Then
            For Each itm In Me.attachments
                .attachments.Add itm
            Next
            
        End If
        .Send
    End With
    
    Set olMail = Nothing
End Sub

Public Sub GetMailDetail(strCriteria As String)
    Dim OTab As Object
    Dim olRow As Object
    Dim olFolder As Object
    Dim OlFolder1 As Object
    Dim ns As Object
    Dim vrPath As Variant
    Dim lngCounter As Long
    Dim olMail As Object
    Dim rst As Recordset
    Dim tblDef As TableDef
    Dim fld As Field
    Dim db As DAO.Database
    Dim Rst2 As DAO.Recordset2
    Dim strPath As String
    
    
    Set ns = olApp.GetNameSpace("MAPI")
    
    If Me.inboxpath = "Inbox" Then
        Set olFolder = olApp.Session.GetDefaultFolder(6)
    Else
        vrPath = Split(Me.inboxpath, "\")
        Set olFolder = ns.Folders(vrPath(LBound(vrPath)))
        For lngCounter = LBound(vrPath) + 1 To UBound(vrPath)
            Set olFolder = olFolder.Folders(vrPath(lngCounter))
        Next
    End If
    Set OTab = olFolder.gettable(strCriteria)
    
    DoCmd.SetWarnings False
    CurrentDb.Execute "drop Table tmpMailFound"
    DoCmd.RunSQL "create table tmpMailFound(Subject  text(255),ReceivedTime Date,From text(255),Body Longtext)"
    Set db = CurrentDb
    Set tblDef = db.TableDefs("tmpMailFound")
    Set fld = tblDef.CreateField("MailAttachments", dbAttachment)
    tblDef.Fields.Append fld
    Set rst = CurrentDb.OpenRecordset("select * from tmpMailFound")
    Do Until (OTab.EndOfTable)
        Set orow = OTab.GetNextRow()
        Set olMail = ns.GetItemFromID(orow("EntryID"))
        If Not olMail Is Nothing Then
            rst.AddNew
            rst![Subject] = olMail.Subject
            rst![ReceivedTime] = olMail.ReceivedTime
            rst![Body] = olMail.Body
            rst![FROM] = olMail.SenderEmailAddress
            For Each atch In olMail.attachments
                strPath = CurrentProject.path & "\" & atch.FileName
                atch.SaveAsFile strPath
                Set Rst2 = rst.Fields("MailAttachments").value
                AppendFile Rst2, strPath
            Next
            rst.Update
        End If
        Set olMail = Nothing
    Loop
End Sub
    
Sub AppendFile(AttachmentRs As Recordset2, strFile As String)
    Dim FileData As Field2
    AttachmentRs.AddNew
    Set FileData = AttachmentRs.Fields(0)
    FileData.LoadFromFile strFile
    AttachmentRs.Update
End Sub


Public Sub SaveAttachmentsToDisk(strPath As String, Optional Subject As String = "", Optional sTo As String = "", Optional SenderName As String = "", Optional SentOnBehalfOfName As String = "", Optional ReplyRecipientNames As String = "", Optional ReceivedByName As String, Optional ConversationTopic As String = "", Optional CC As String, Optional BCC As String, Optional Body As String)
     Dim OTab As Object
    Dim olRow As Object
    Dim olFolder As Object
    Dim OlFolder1 As Object
    Dim ns As Object
    Dim vrPath As Variant
    Dim lngCounter As Long
    Dim olMail As Object
    
    If Subject <> "" Then
        
    End If
    
    
    
    Set ns = olApp.GetNameSpace("MAPI")
    
    If Me.inboxpath = "Inbox" Then
        Set olFolder = olApp.Session.GetDefaultFolder(6)
    Else
        vrPath = Split(Me.inboxpath, "\")
        Set olFolder = ns.Folders(vrPath(LBound(vrPath)))
        For lngCounter = LBound(vrPath) + 1 To UBound(vrPath)
            Set olFolder = olFolder.Folders(vrPath(lngCounter))
        Next
    End If
    Set OTab = olFolder.gettable(strCriteria)
    
    Do Until (OTab.EndOfTable)
        Set orow = OTab.GetNextRow()
        Set olMail = ns.GetItemFromID(orow("EntryID"))
        If Not olMail Is Nothing Then
            For Each atch In olMail.attachments
                atch.SaveAsFile strPath & "\" & atch.FileName
            Next
        End If
        Set olMail = Nothing
    Loop
End Sub
