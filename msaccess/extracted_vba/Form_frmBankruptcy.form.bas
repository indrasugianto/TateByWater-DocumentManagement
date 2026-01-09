' Component: Form_frmBankruptcy
' Type: document
' Lines: 174
' ============================================================

Option Compare Database

Private Sub cmbtrustee_AfterUpdate()
    If Me.cmbtrustee = "Thomas P. Gorman" Then
        Me.TrusteeAddress = "300 N. Washington Street, Ste. 400"
        Me.TrusteeCity = "Alexandria"
        Me.TrusteeState = "VA"
        Me.TrusteeZip = "22314"
        Me.TrusteePhone = "703-836-2226"
        Me.TrusteeFax = "703-836-8120"
        Me.TrusteeEmail = "TGorman@chapter13alexva.com"
    End If
      If Me.cmbtrustee = "H. Jason Gold" Then
        Me.TrusteeAddress = "101 Constitution Avenue, NW, Ste. 900"
        Me.TrusteeCity = "Washington"
        Me.TrusteeState = "DC"
        Me.TrusteeZip = "20001"
        Me.TrusteePhone = "202-689-2819"
        Me.TrusteeFax = "202-712-2860"
        Me.TrusteeEmail = "jason.gold@nelsonmullins.com"
    End If
      If Me.cmbtrustee = "Donald F. King" Then
        Me.TrusteeAddress = "1775 Wiehle Avenue, Ste. 400"
        Me.TrusteeCity = "Reston"
        Me.TrusteeState = "VA"
        Me.TrusteeZip = "20190"
        Me.TrusteePhone = "703-218-2116"
        Me.TrusteeFax = "703-218-2160"
        Me.TrusteeEmail = "DonKing@ofplaw.com"
    End If
      If Me.cmbtrustee = "Kevin R. McCarthy" Then
        Me.TrusteeAddress = "1751 Pinnacle Drive, Ste. 1115"
        Me.TrusteeCity = "McLean"
        Me.TrusteeState = "VA"
        Me.TrusteeZip = "22102"
        Me.TrusteePhone = "703-770-9261"
        Me.TrusteeFax = "703-770-9264"
        Me.TrusteeEmail = "krm@mccarthywhite.com"
    End If
      If Me.cmbtrustee = "Janet M. Meiburger" Then
        Me.TrusteeAddress = "1493 Chain Bridge Road, Ste. 201"
        Me.TrusteeCity = "McLean"
        Me.TrusteeState = "VA"
        Me.TrusteeZip = "22101"
        Me.TrusteePhone = "703-556-7871"
        Me.TrusteeFax = "703-556-8609"
        Me.TrusteeEmail = "janetm@meiburgerlaw.com"
    End If
      If Me.cmbtrustee = "Bruce H. Matson" Then
        Me.TrusteeAddress = "919 E. Main Street, 24th Floor"
        Me.TrusteeCity = "Richmond"
        Me.TrusteeState = "VA"
        Me.TrusteeZip = "23219"
        Me.TrusteePhone = "804-783-2003"
        Me.TrusteeFax = "804-783-7629"
        Me.TrusteeEmail = "bruce.matson@leclairryan.com"
    End If
        
End Sub

Private Sub cmdPrintForeLabel_Click()
    Me.Refresh
 On Error GoTo ErrHandler_cmdPrintForeLabel_Click

    answer = MsgBox("Print Address Label?", vbYesNo, "TB CMS Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_ftrustee_address_Label").IsLoaded Then
            DoCmd.Close acReport, "rpt_ftrustee_address_Label", acSaveNo
        End If
        DoCmd.OpenReport "rpt_ftrustee_address_Label", acNormal, , "[CaseID]=" & CaseID
    End If
    
ErrHandler_cmdPrintForeLabel_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub cmdPrintTrusteeLabel_Click()
Me.Refresh
On Error GoTo ErrHandler_cmdPrintTrusteeLabel_Click

    answer = MsgBox("Print Address Label?", vbYesNo, "TB CMS: Direct Print")
    If answer = vbYes Then
        If CurrentProject.AllReports("rpt_trustee_address_Label").IsLoaded Then
            DoCmd.Close acReport, "rpt_trustee_address_Label", acSaveNo
        End If
        DoCmd.OpenReport "rpt_trustee_address_Label", acNormal, , "[CaseID]=" & CaseID
    End If
    
ErrHandler_cmdPrintTrusteeLabel_Click:
    If Err.Number <> 0 Then ShowMessage Err.Description
End Sub

Private Sub Fore_Trustee_AfterUpdate()
    If Me.ForeTrustee = "Atlantic Law Group / Orlans PC" Then
        Me.ForeAddress = "1602 Village Market Boulevard SE, Ste. 310"
        Me.ForeCity = "Leesburg"
        Me.ForeState = "VA"
        Me.ForeZIP = "20175"
        Me.ForePhone = "703-777-7101"
        Me.ForeFax = "703-940-9111"
    End If
    If Me.ForeTrustee = "Brock & Scott, PLLC" Then
        Me.ForeAddress = "484 Viking Drive, Ste. 203"
        Me.ForeCity = "Virginia Beach"
        Me.ForeState = "VA"
        Me.ForeZIP = "23452"
        Me.ForePhone = "757-213-2959"
        Me.ForeFax = "703-840-4279"
    End If
    If Me.ForeTrustee = "BWW Law Group, LLC" Then
        Me.ForeAddress = "8100 Three Chopt Road, Ste. 240"
        Me.ForeCity = "Richmond"
        Me.ForeState = "VA"
        Me.ForeZIP = "23229"
        Me.ForePhone = "804-282-0463"
        Me.ForeFax = "804-282-0541"
    End If
    If Me.ForeTrustee = "Commonwealth Trustees, LLC" Then
        Me.ForeAddress = "8601 Westwood Center Drive, Ste. 255"
        Me.ForeCity = "Woodbridge"
        Me.ForeState = "VA"
        Me.ForeZIP = "22182"
        Me.ForePhone = "703-752-8500"
        Me.ForeFax = "703-752-4300"
    End If
    If Me.ForeTrustee = "Glasser and Glasser, PLC" Then
        Me.ForeAddress = "580 East Main Street, Ste. 600"
        Me.ForeCity = "Norfolk"
        Me.ForeState = "VA"
        Me.ForeZIP = "23510-2212"
        Me.ForePhone = "757-625-6787"
        Me.ForeFax = "757-625-5959"
    End If
    If Me.ForeTrustee = "Samuel I. White, P.C." Then
        Me.ForeAddress = "5040 Corporate Woods Drive, Ste. 120"
        Me.ForeCity = "Virginia Beach"
        Me.ForeState = "VA"
        Me.ForeZIP = "23462"
        Me.ForePhone = "757-490-9284"
        Me.ForeFax = "757-497-2802"
    End If
    If Me.ForeTrustee = "Shapiro & Brown, LLP" Then
        Me.ForeAddress = "10021 Balls Ford Road, Ste. 200"
        Me.ForeCity = "Manassas"
        Me.ForeState = "VA"
        Me.ForeZIP = "20109"
        Me.ForePhone = "703-449-5800"
        Me.ForeFax = "703-449-5850"
    End If
    If Me.ForeTrustee = "The O'Reilly Law Firm" Then
        Me.ForeAddress = "761-C Monroe Street, Ste. 200"
        Me.ForeCity = "Herndon"
        Me.ForeState = "VA"
        Me.ForeZIP = "20170"
        Me.ForePhone = "703-766-1991"
        Me.ForeFax = "703-766-1995"
    End If
    If Me.ForeTrustee = "Rosenberg & Associates, LLC" Then
        Me.ForeAddress = "8601 Westwood Center Drive, Ste. 255"
        Me.ForeCity = "Vienna"
        Me.ForeState = "VA"
        Me.ForeZIP = "22182"
        Me.ForePhone = "301-907-8000"
        Me.ForeFax = "301-907-8101"
    End If
    If Me.ForeTrustee = "Surety Trustees, LLC" Then
        Me.ForeAddress = "722 East Market Street, Ste. 203"
        Me.ForeCity = "Leesburg"
        Me.ForeState = "VA"
        Me.ForeZIP = "20176"
        Me.ForePhone = "571-449-9350"
        Me.ForeFax = "855-845-2585"
    End If
End Sub