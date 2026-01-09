' Component: ModGeneric Func
' Type: module
' Lines: 48
' ============================================================

Option Compare Database

Public Function ValidateFields(TargetForm As Form, MandaToryFields As Variant) As Boolean
    Dim nCounter As Long
    Dim TextBx As TextBox
    Dim cmb As ComboBox
    
    For nCounter = LBound(MandaToryFields) To UBound(MandaToryFields)
        If TargetForm.Controls(MandaToryFields(nCounter)).ControlType = acTextBox Then
            Set TextBx = TargetForm.Controls(MandaToryFields(nCounter))
            If Nz(TextBx.value, "") = "" Then
                ShowMessage "Please enter " & TextBx.Name
                TextBx.SetFocus
                ValidateFields = False
                Exit Function
            End If
        ElseIf TargetForm.Controls(MandaToryFields(nCounter)).ControlType = acComboBox Then
            Set cmb = TargetForm.Controls(MandaToryFields(nCounter))
            If Nz(cmb.value, "") = "" Then
                ShowMessage "Please enter " & cmb.Name
                cmb.SetFocus
                ValidateFields = False
                Exit Function
            End If
        End If
        
        Set TextBx = Nothing
        Set cmb = Nothing
    Next
    
    ValidateFields = True
End Function


Sub TestOl()
    Dim ol As OutlookApp
    Set ol = New OutlookApp
    
    
    ol.CreateApp
    
  'ol.AddAttachments "C:\Users\EADC0007\Desktop\Attachments\aaa.jpg"
  'ol.AddAttachments "C:\Users\EADC0007\Desktop\Attachments\DB-NEWMARK.PNG"
   'ol.SendMail "Test", "Testing", "arup.banerjee@oscillateinfo.co.in"

    ol.inboxpath = "arup.banerjee@oscillateinfo.co.in\Inbox"
    ol.GetMailDetail "[ReceivedByName]='Arup Banerjee'"
End Sub