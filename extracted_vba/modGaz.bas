' Component: modGaz
' Type: module
' Lines: 418
' ============================================================

Option Compare Database

Function fncRunningDebit(intCaseID As Integer, dtDate As Date, MatterID As Integer)

    'this function is specific to [Matter and AR] table

    'Debug.Print DSum("Charge", "[Matter and AR]", "CaseID=11 AND Date2<=#" & "09/15/16" & "#")
    strQuery = "CaseID=" & intCaseID & _
               " AND MatterID<=" & MatterID & _
               " AND Date2<=#" & dtDate & "#"
    retval = DSum("Charge", "[Matter and AR]", strQuery)
    If retval > 0 Then
        fncRunningDebit = retval
    Else
        fncRunningDebit = 0
    End If
End Function

Function fncRunningCredit(intCaseID As Integer, dtDate As Date, MatterID As Integer)

    'this function is specific to [Matter and AR] table

    'Debug.Print DSum("Payment", "[Matter and AR]", "CaseID=11 AND Date2<=#" & "09/15/16" & "#")
    strQuery = "CaseID=" & intCaseID & _
               " AND MatterID<=" & MatterID & _
               " AND Date2<=#" & dtDate & "#"
    retval = DSum("Payment", "[Matter and AR]", strQuery)
    If retval > 0 Then
        fncRunningCredit = retval
    Else
        fncRunningCredit = 0
    End If
End Function

Function fncRunningDebitStmtTrust(intCaseID As Integer, dtDate As Date, TrustAccountID As Integer)
    
    'this function is specific to [Trust Account] table

    strQuery = "CaseID=" & intCaseID & _
               " AND TrustAccountID<=" & TrustAccountID & _
               " AND TDate<=#" & dtDate & "#"
    retval = DSum("Debit", "[Trust Account]", strQuery)
    If retval > 0 Then
        fncRunningDebitStmtTrust = retval
    Else
        fncRunningDebitStmtTrust = 0
    End If
End Function

Function fncRunningCreditStmtTrust(intCaseID As Integer, dtDate As Date, TrustAccountID As Integer)
    
    'this function is specific to [Trust Account] table
    
    strQuery = "CaseID=" & intCaseID & _
               " AND TrustAccountID<=" & TrustAccountID & _
               " AND TDate<=#" & dtDate & "#"
    retval = DSum("Credit", "[Trust Account]", strQuery)
    If retval > 0 Then
        fncRunningCreditStmtTrust = retval
    Else
        fncRunningCreditStmtTrust = 0
    End If
End Function

Function fncGetLastInvoiceSent(intCaseID As Integer)
        
    retval = DLookup("LastOfInvSent", "qry_last_invoice_sent", "CaseID=" & intCaseID)
        
    If retval > 0 Then
        fncGetLastInvoiceSent = retval
    Else
        fncGetLastInvoiceSent = Null
    End If
End Function

Function fncTakeOffSums(FieldName As String, intTakeOffMonthID As Integer)
    'On Error Resume Next
    'MsgBox intTakeOffMonthID
    'Working: =fncTakeOffSums([Forms]![frmTakeOff]![TakeOffMonthID])
    
    'NOT WORKING:  =DLookup("SumOfCBHRev", "qry_take_off_step2_sums", [Forms]![frmTakeOff]![TakeOffMonthID])
    
    'now trying for field name variable:
    'fncTakeOffSums = DLookup("SumOfCBHRev", "qry_take_off_step2_sums", intTakeOffMonthID)
    result = DLookup(FieldName, "qry_take_off_step2_sums", "TakeOffMonthID=" & intTakeOffMonthID)
    'Debug.Print result
    fncTakeOffSums = Nz(result, 0)
    
End Function

Function fncTakeOffSumsAttorneys(intTakeOffMonthID As Integer, strOrig_Atty As String, strCase_Letter As String, strFieldName As String) As Currency
On Error GoTo ERR_HANDLER
Dim retValue As Currency

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String

    'Examples:
    '
    'fncTakeOffSumsAttorneys(142, "PM", "T", "SumOfEarned")
    'fncTakeOffSumsAttorneys(142, "PM", "T", "SumOfBilled")
    'fncTakeOffSumsAttorneys(142, "PM", "T", "SumOfAdvFeeBal")
    'pass "%" to Case_Letter to calculate all case types
    'fncTakeOffSumsAttorneys(142, "PM", "%", "SumOfEarned")

    Set cn = New ADODB.Connection
    cn.Open PcaGetConnnectionString()
    
    sql = ""
    sql = sql & "exec spGetTakeOffSumsAttorneys "
    sql = sql & "@TakeOffMonthID = " & intTakeOffMonthID
    sql = sql & ",@Orig_Atty = " & pcaAddQuotes(strOrig_Atty)
    sql = sql & ",@Case_Letter = " & pcaAddQuotes(strCase_Letter)
    
    Set rs = New ADODB.Recordset
    Set rs = cn.Execute(sql)
    
    If Not rs.EOF Then
        rs.MoveFirst
        retValue = pcaConvertNulls(rs(strFieldName), 0)
    End If
EXIT_HANDLER:
    fncTakeOffSumsAttorneys = retValue
    Exit Function
ERR_HANDLER:
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
End Function


Function fncReturnNullOnError(strField As String)
    MsgBox "aaa"
    Debug.Print strField
    fncReturnNullOnError = Nz(strField, "0")
End Function

Function fncGetTrustAccountBalance(OrderNr As Integer)
    Dim CaseID As Long

    CaseID = GetCaseID()

    strWhere = "OrderNr<=" & OrderNr & " and CaseID=" & CaseID '& " order by OrderNr asc"
    Debug.Print strWhere
    debitVal = DSum("SumOfDebit", "[qryTrustAccount]", strWhere)
    creditVal = DSum("SumOfCredit", "[qryTrustAccount]", strWhere)

    balanceVal = debitVal - creditVal
    fncGetTrustAccountBalance = balanceVal
End Function

'Function fncGetMatterARBalance(OrderNr As Integer)
'    Dim CaseID As Long
'
'    CaseID = GetCaseID()
'    retainerVal = GetRetainer(CaseID)
'
'    'strWhere = "Date2<=#" & Date2 & "# and MatterID<=" & MatterID & " and CaseID=" & CaseID
'    strWhere = "OrderNr<=" & OrderNr & " and CaseID=" & CaseID '& " order by OrderNr asc"
'    'Debug.Print strWhere
'    chargeVal = DSum("SumOfCharge", "[qryMatter]", strWhere)
'    paymentVal = DSum("SumOfPayment", "[qryMatter]", strWhere)
'
'    balanceVal = retainerVal + chargeVal - paymentVal
'    fncGetMatterARBalance = balanceVal
'End Function

Function fncGetTABalanceWithCaseID(OrderNr As Long, CaseID As Long)
    'retainerVal = DLookup("Retainer", "tblCase", "CaseID=" & CaseID)
    
    'strWhere = "Date2<=#" & Date2 & "# and MatterID<=" & MatterID & " and CaseID=" & CaseID
    strWhere = "OrderNr<=" & OrderNr & " and CaseID=" & CaseID '& " order by OrderNr asc"
    'Debug.Print strWhere
    chargeVal = DSum("SumOfDebit", "[qryTrustAccount]", strWhere)
    paymentVal = DSum("SumOfCredit", "[qryTrustAccount]", strWhere)
    
    balanceVal = chargeVal - paymentVal
    fncGetTABalanceWithCaseID = balanceVal
End Function

Function fncGetFilterOrderNrTA(CaseID As Long)
    intCaseID = Nz(Form_frmClientLedger.CaseID, 0)
    strFilter = "CaseID=" & intCaseID & " AND Balance=0"
    intOrderNr = DMax("OrderNr", "qry_invoice_comprehensive_trust_acc_cur_unfiltered", strFilter)
    
    If IsNull(intOrderNr) Then
        strOrderNr = ""
    Else
        strOrderNr = "[OrderNr] > " & intOrderNr
    End If
    Debug.Print "Filter is:" & strOrderNr
    fncGetFilterOrderNrTA = strOrderNr
End Function

Function fncGetFilterOrderNrMatterAR(CaseID As Long)
    intCaseID = Nz(Form_frmClientLedger.CaseID, 0)
    strFilter = "CaseID=" & intCaseID & " AND Balance=0"
    intOrderNr = DMax("OrderNr", "qry_current_invoice", strFilter)
    
    If IsNull(intOrderNr) Then
        strOrderNr = ""
    Else
        strOrderNr = "[OrderNr] > " & intOrderNr
    End If
    'Debug.Print "Filter is:" & strOrderNr
    fncGetFilterOrderNrMatterAR = strOrderNr
End Function

Function English(ByVal n As Currency) As String

   Const Thousand = 1000@
   Const Million = Thousand * Thousand
   Const Billion = Thousand * Million
   Const Trillion = Thousand * Billion

   If (n = 0@) Then English = "zero": Exit Function

   Dim Buf As String: If (n < 0@) Then Buf = "negative " Else Buf = ""
   Dim Frac As Currency: Frac = Abs(n - Fix(n))
   If (n < 0@ Or Frac <> 0@) Then n = Abs(Fix(n))
   Dim AtLeastOne As Integer: AtLeastOne = n >= 1

   If (n >= Trillion) Then
      Debug.Print n
      Buf = Buf & EnglishDigitGroup(Int(n / Trillion)) & " trillion"
      n = n - Int(n / Trillion) * Trillion
      If (n >= 1@) Then Buf = Buf & " "
   End If

   If (n >= Billion) Then
      Debug.Print n
      Buf = Buf & EnglishDigitGroup(Int(n / Billion)) & " billion"
      n = n - Int(n / Billion) * Billion
      If (n >= 1@) Then Buf = Buf & " "
   End If

   If (n >= Million) Then
      Debug.Print n
      Buf = Buf & EnglishDigitGroup(n \ Million) & " million"
      n = n Mod Million
      If (n >= 1@) Then Buf = Buf & " "
   End If

   If (n >= Thousand) Then
      Debug.Print n
      Buf = Buf & EnglishDigitGroup(n \ Thousand) & " thousand"
      n = n Mod Thousand
      If (n >= 1@) Then Buf = Buf & " "
   End If

   If (n >= 1@) Then
      Debug.Print n
      Buf = Buf & EnglishDigitGroup(n)
   End If

   If (Frac = 0@) Then
      Buf = Buf & " exactly"
   ElseIf (Int(Frac * 100@) = Frac * 100@) Then
      If AtLeastOne Then Buf = Buf & " and "
      Buf = Buf & Format$(Frac * 100@, "00") & "/100"
   Else
      If AtLeastOne Then Buf = Buf & " and "
      Buf = Buf & Format$(Frac * 10000@, "0000") & "/10000"
   End If

   English = Buf
End Function

Private Function EnglishDigitGroup(ByVal n As Integer) As String

   Const Hundred = " hundred"
   Const One = "one"
   Const Two = "two"
   Const Three = "three"
   Const Four = "four"
   Const Five = "five"
   Const Six = "six"
   Const Seven = "seven"
   Const Eight = "eight"
   Const Nine = "nine"
   Dim Buf As String: Buf = ""
   Dim Flag As Integer: Flag = False

   Select Case (n \ 100)
      Case 0: Buf = "": Flag = False
      Case 1: Buf = One & Hundred: Flag = True
      Case 2: Buf = Two & Hundred: Flag = True
      Case 3: Buf = Three & Hundred: Flag = True
      Case 4: Buf = Four & Hundred: Flag = True
      Case 5: Buf = Five & Hundred: Flag = True
      Case 6: Buf = Six & Hundred: Flag = True
      Case 7: Buf = Seven & Hundred: Flag = True
      Case 8: Buf = Eight & Hundred: Flag = True
      Case 9: Buf = Nine & Hundred: Flag = True
   End Select

   If (Flag <> False) Then n = n Mod 100
   If (n > 0) Then
      If (Flag <> False) Then Buf = Buf & " "
   Else
      EnglishDigitGroup = Buf
      Exit Function
   End If

   Select Case (n \ 10)
      Case 0, 1: Flag = False
      Case 2: Buf = Buf & "twenty": Flag = True
      Case 3: Buf = Buf & "thirty": Flag = True
      Case 4: Buf = Buf & "forty": Flag = True
      Case 5: Buf = Buf & "fifty": Flag = True
      Case 6: Buf = Buf & "sixty": Flag = True
      Case 7: Buf = Buf & "seventy": Flag = True
      Case 8: Buf = Buf & "eighty": Flag = True
      Case 9: Buf = Buf & "ninety": Flag = True
   End Select

   If (Flag <> False) Then n = n Mod 10
   If (n > 0) Then
      If (Flag <> False) Then Buf = Buf & "-"
   Else
      EnglishDigitGroup = Buf
      Exit Function
   End If

   Select Case (n)
      Case 0:
      Case 1: Buf = Buf & One
      Case 2: Buf = Buf & Two
      Case 3: Buf = Buf & Three
      Case 4: Buf = Buf & Four
      Case 5: Buf = Buf & Five
      Case 6: Buf = Buf & Six
      Case 7: Buf = Buf & Seven
      Case 8: Buf = Buf & Eight
      Case 9: Buf = Buf & Nine
      Case 10: Buf = Buf & "ten"
      Case 11: Buf = Buf & "eleven"
      Case 12: Buf = Buf & "twelve"
      Case 13: Buf = Buf & "thirteen"
      Case 14: Buf = Buf & "fourteen"
      Case 15: Buf = Buf & "fifteen"
      Case 16: Buf = Buf & "sixteen"
      Case 17: Buf = Buf & "seventeen"
      Case 18: Buf = Buf & "eighteen"
      Case 19: Buf = Buf & "nineteen"
   End Select

   EnglishDigitGroup = Buf

End Function

Function get_remaining_AdvancedChargesBalance(lngCaseID As Long, lngMatterID As Long) As Currency

    If CaseID = 9966 Then Stop
    
    Dim rs As Recordset
    
    strSQL = "select * from [qry_advanced_nonadvanced_payments] where CaseID = " & lngCaseID & " and MatterID < " & lngMatterID 'Form_frmClientLedger.CaseID
    'strSQL = "select * from [qry_advanced_nonadvanced_payments_01]"
    'Debug.Print strSQL
    
    lngRetainer = DLookup("Retainer", "tblCase", "CaseID = " & lngCaseID)
        
    strRows = lngRetainer & ";0;0;" & lngRetainer & vbCrLf
    
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
    
    i = 0
    Do Until rs.EOF
        pre_balance = lngRetainer + pre_balance + rs("NonAdvancedCharges") + rs("AdvancedCharges") - rs("PaymentMade") '(see Excel File - formula at column L)
        strRows = strRows & rs("NonAdvancedCharges") & ";" & rs("AdvancedCharges") & ";" & rs("PaymentMade") & ";" & pre_balance & vbCrLf
            
        lngRetainer = 0
        i = i + 1
        rs.MoveNext
    Loop
    
    strRows = Left(strRows, Len(strRows) - 1)
    
    oneRow = Split(strRows, vbCrLf)
        
    For p = 0 To UBound(oneRow) - 1
        'Debug.Print oneRow(p)
        
        i = Split(oneRow(p), ";")(0) 'NonAdvancedCharges
        j = Split(oneRow(p + 1), ";")(1) 'AdvancedCharges '!Next row (see Excel File - formula at column M)
        k = Split(oneRow(p + 1), ";")(2) 'PaymentMade '!Next row (see Excel File - formula at column M)
        l = Split(oneRow(p), ";")(3) 'pre_balance
        If i = "" Then i = 0
        If j = "" Then j = 0
        If k = "" Then k = 0
        If l = "" Then l = 0
        lngBalance = -i + j - k + l 'balance
        'Debug.Print lngBalance
    Next
    If lngBalance < 0 Then lngBalance = 0
    get_remaining_AdvancedChargesBalance = lngBalance
End Function

Sub fnc_TEST_get_remaining_AdvancedChargesBalance()
    lngPayment = DLookup("Payment", "Matter And AR", "MatterID = " & 2265)
    Debug.Print lngPayment
    Debug.Print get_remaining_AdvancedChargesBalance(9966, 2265)
    Debug.Print lngPayment - get_remaining_AdvancedChargesBalance(9966, 2265)
End Sub

Function fncGetMatterARBalanceWithCaseID(OrderNr As Long, CaseID As Long)
    retainerVal = DLookup("Retainer", "tblCase", "CaseID=" & CaseID)
    
    'strWhere = "Date2<=#" & Date2 & "# and MatterID<=" & MatterID & " and CaseID=" & CaseID
    strWhere = "OrderNr<=" & OrderNr & " and CaseID=" & CaseID '& " order by OrderNr asc"
    'Debug.Print strWhere
    chargeVal = DSum("SumOfCharge", "[qryMatter]", strWhere)
    paymentVal = DSum("SumOfPayment", "[qryMatter]", strWhere)
    
    balanceVal = retainerVal + chargeVal - paymentVal
    fncGetMatterARBalanceWithCaseID = balanceVal
End Function