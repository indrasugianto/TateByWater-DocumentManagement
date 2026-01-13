' Component: PcaStdLib
' Type: module
' Lines: 1947
' ============================================================

Option Explicit
Option Compare Text
Public Const MAX_PATH = 260
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'Declare Function SQLWriteDSNToIni Lib "odbccp32.dll" (ByVal lpszDSN As String, ByVal lpszDriver As String) As Boolean
'Declare Function SQLDataSources Lib "odbc32.dll" ()
Declare PtrSafe Function SQLRemoveDSNFromIni Lib "ODBCCP32.DLL" (ByVal lpszDSN As String) As Boolean
Declare PtrSafe Function SQLWritePrivateProfileString Lib "ODBCCP32.DLL" (ByVal lpszSection As String, ByVal lpszEntry As String, ByVal lpszString As String, ByVal lpszFilename As String) As Boolean
Private Declare PtrSafe Function GetSystemDirectoryA Lib "kernel32" (ByVal lpszDirName As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpszDirName As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function APIGetParent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function APIgetClientRect Lib "user32" Alias "GetClientRect" (ByVal hwnd As Long, lprect As RECT) As Long
Private Declare PtrSafe Function APIisZoomed Lib "user32" Alias "IsZoomed" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function APIGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function APIisWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function APIEnableWindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function APIupdateWindow Lib "user32" Alias "UpdateWindow" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function APIMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'// Api calls to read .ini file
'Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private pcaFileSelect_RecentFileList As New Collection

Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Declare PtrSafe Function aht_apiGetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean

Declare PtrSafe Function aht_apiGetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

'------------------------------------------------------------------------
' GLOBAL CONST FOR MS DIALOG
'------------------------------------------------------------------------
Global Const ahtOFN_READONLY = &H1
Global Const ahtOFN_OVERWRITEPROMPT = &H2
Global Const ahtOFN_HIDEREADONLY = &H4
Global Const ahtOFN_NOCHANGEDIR = &H8
Global Const ahtOFN_SHOWHELP = &H10
' You won't use these.
'Global Const ahtOFN_ENABLEHOOK = &H20
'Global Const ahtOFN_ENABLETEMPLATE = &H40
'Global Const ahtOFN_ENABLETEMPLATEHANDLE = &H80
Global Const ahtOFN_NOVALIDATE = &H100
Global Const ahtOFN_ALLOWMULTISELECT = &H200
Global Const ahtOFN_EXTENSIONDIFFERENT = &H400
Global Const ahtOFN_PATHMUSTEXIST = &H800
Global Const ahtOFN_FILEMUSTEXIST = &H1000
Global Const ahtOFN_CREATEPROMPT = &H2000
Global Const ahtOFN_SHAREAWARE = &H4000
Global Const ahtOFN_NOREADONLYRETURN = &H8000
Global Const ahtOFN_NOTESTFILECREATE = &H10000
Global Const ahtOFN_NONETWORKBUTTON = &H20000
Global Const ahtOFN_NOLONGNAMES = &H40000
' New for Windows 95
Global Const ahtOFN_EXPLORER = &H80000
Global Const ahtOFN_NODEREFERENCELINKS = &H100000
Global Const ahtOFN_LONGNAMES = &H200000

Dim foo

Function GetWindowsDirectory()
Dim strDir As String, nBytes As Integer
nBytes = 255
strDir = Space(nBytes)
nBytes = GetWindowsDirectoryA(strDir, nBytes)
GetWindowsDirectory = Mid(strDir, 1, nBytes)
End Function
Function GetSystemDirectory()
Dim strDir As String, nBytes As Integer
nBytes = 255
strDir = Space(nBytes)
nBytes = GetSystemDirectoryA(strDir, nBytes)
GetSystemDirectory = Mid(strDir, 1, nBytes)
End Function

Function FormRestoreFullScreen(f As Form)
Dim MaxHeight As Long, MaxWidth As Long
Dim nTop As Long, nLeft As Long

foo = pcaGetClientRect(f.hwnd, nLeft, nTop, MaxWidth, MaxHeight)
foo = FormGetParentClientRect(f, nLeft, nTop, MaxWidth, MaxHeight)
foo = pcaMoveWindow(f.hwnd, nLeft, nTop, MaxWidth - 5, MaxHeight - 5, True)

End Function
Function FormGetParentClientRect(f As Form, fLeft As Long, fTop As Long, fWidth As Long, fHeight As Long)
foo = pcaGetClientRect(pcaGetParent(f.hwnd), fLeft, fTop, fWidth, fHeight)
End Function
Public Function pcaGetClientRect(hwnd As Long, fLeft As Long, fTop As Long, fWidth As Long, fHeight As Long)
Dim R As RECT
R.Bottom = fHeight
R.Left = fLeft
R.Right = fWidth
R.Top = fTop
pcaGetClientRect = APIgetClientRect(hwnd, R)
fLeft = R.Left '- rectParent.x1
fTop = R.Top '- rectParent.y1
fWidth = R.Right '- rectParent.x1 - fLeft
fHeight = R.Bottom '- rectParent.y1 - fTop
End Function

Public Function pcaGetParent(hwnd As Long) As Long
pcaGetParent = APIGetParent(hwnd)
End Function

Public Function pcaMoveWindow(hwnd As Long, X As Long, Y As Long, nWidth As Long, nHeight As Long, bRepaint As Long) As Long
pcaMoveWindow = APIMoveWindow(hwnd, X, Y, nWidth, nHeight, bRepaint)
End Function

Public Function pcaUpdateWindow(hwnd As Long) As Long
pcaUpdateWindow = APIupdateWindow(hwnd)
End Function
 
Public Function pcaEnableWindow(hwnd As Long, lEnable As Long) As Long
pcaEnableWindow = APIEnableWindow(hwnd, lEnable)
End Function

Public Function pcaIsWindowEnabled(hwnd As Long) As Long
pcaIsWindowEnabled = APIisWindowEnabled(hwnd)
End Function
Public Function pcaisZoomed(hwnd As Long) As Long
pcaisZoomed = APIisZoomed(hwnd)
End Function
Public Function pcaGetClassName(hwnd As Long) As String
    Dim cch As Long
    Const cchMax = 255
    Dim stBuff As String * cchMax
    cch = APIGetClassName(hwnd, stBuff, cchMax)
    If (hwnd = 0) Then
        pcaGetClassName = ""
    Else
        pcaGetClassName = (Left$(stBuff, cch))
    End If
End Function

Public Function pcaAddPounds(d As Variant) As String
pcaAddPounds = "#" & CVDate(d) & "#"
End Function
Public Function pcaAddQuotes(s As Variant)
On Error GoTo ERR_AddQuotes
Dim foo
Select Case VarType(s)
Case 0, 1
    pcaAddQuotes = " Null "
Case 8, 2, 3, 4, 5, 6
    pcaAddQuotes = "'" & pcaConvertSingleQuote(CStr(s)) & "'"
'Case 2, 3, 4, 5, 6  ' this doesn't work for controls which are string values with all numeric characters.
'    Addquotes = s
Case 7
    'AddQuotes = AddPounds(CVDate(s)) ' This doesn't work for sqlserver or oracle which expect strings.
    pcaAddQuotes = "'" & CVDate(s) & "'"
End Select
Exit Function
'-------------
ERR_AddQuotes:
Select Case Err
Case 2427
    Resume Next
End Select
foo = MsgBox(Err, Error)
Resume Next
Resume
End Function
Public Function pcaConvertSingleQuote(s As String)
Dim POS As Integer
POS = InStr(s, "'")
If POS Then
    pcaConvertSingleQuote = Mid$(s, 1, POS) & "'" & pcaConvertSingleQuote(Mid$(s, POS + 1))
Else
    pcaConvertSingleQuote = s
End If

End Function

Function pcaCrlf() As String
pcaCrlf = vbCrLf
End Function

Public Function pcaReplaceStr(ByVal stSource As String, stWhat As String, stWith As String) As String
Dim i As Integer
Do While True
    i = InStr(stSource, stWhat)
    If i > 0 Then
        stSource = Left$(stSource, i - 1) & stWith & Right(stSource, Len(stSource) - Len(stWhat) - i + 1)
    Else
        Exit Do
    End If
Loop
pcaReplaceStr = stSource
End Function

Public Function pcaStrExtract(ByVal s As String, c As String, n As Integer) As String
Dim POS As Integer, i As Integer, nTemp  As Integer
i = 1
s = s & c
Do
   If i = n Then ' we've found the proper n'th token
      s = Mid$(s, POS + 1)
      nTemp = InStr(1, s, c)
      If nTemp > 0 Then
         pcaStrExtract = Left$(s, nTemp - 1)
      End If

      Exit Function
   End If
   POS = InStr(POS + 1, s, c, 0)
   If POS Then
      i = i + 1
   End If
Loop While POS > 0
      
End Function
Public Function pcaChrCount(c As String, s As String) As Integer
Dim POS As Integer, i As Integer
'pos = 0
'i = 0
Do
   POS = InStr(POS + 1, s, c, 0)
   If POS Then
      i = i + 1
   End If
Loop While POS > 0
pcaChrCount = i
End Function
Function pcaempty(v) As Integer
Dim foo
On Error GoTo Err_pcaEmpty
pcaempty = True
Select Case VarType(v)
Case vbNull, vbEmpty    '(no valid data), Empty (uninitialized)
    pcaempty = True
Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, vbDecimal   'Integer,Long integer,Single,Double, Currency
    pcaempty = (v = 0)
Case vbDate             'Date value
    pcaempty = Not IsDate(v)
Case vbString           'String
    pcaempty = (Trim$(v) = "")
Case vbObject           'Object
    foo = v
    pcaempty = (v Is Nothing)
Case vbError            'Error value
Case vbBoolean   'Boolean value
    pcaempty = Not v
Case vbVariant   'Variant (used only with arrays of variants)
    pcaempty = IsNull(v)
Case vbDataObject 'A data access object
    pcaempty = (v Is Nothing)
'Case vbDecimal   'Decimal value
'Case vbByte  'Byte value
Case vbArray 'Array
Case Else
    MsgBox ("Unrecognized VarType (" & VarType(v) & ") in pcaempty...")
End Select
'----------------------------------------
EXIT_pcaEmpty:
Exit Function
'-----------------------------------------
Err_pcaEmpty:
Select Case Err.Number
Case 2055
    Resume Next
Case 2424   ' Unknown Name
    Resume Next
Case 2467   '       Object referred to in expression no longer exists.
    Resume Next
Case 2427 'You entered an expression that has no value.
'@The expression may refer to an object that has no value, such as a form, a report, or a label control.
    If VarType(v) = vbObject Then
        pcaempty = -1
    GoTo EXIT_pcaEmpty
    End If
    Resume Next
Case 3021 ' No current record
    Resume Next
Case 94
    Resume Next
Case 3167   ' Record is deleted
    Resume Next
End Select
foo = MsgBox(Err.Number, Err.Description)
Resume EXIT_pcaEmpty:
Resume

End Function
Function pcaParseDSNString(ByVal DSN, key)
DSN = pcaReplaceStr(DSN, key & " =", key & "=")  ' remove extraneous spaces
DSN = Mid(DSN, InStr(DSN, key & "=") + Len(key) + 1)
If InStr(DSN, ";") Then DSN = Mid(DSN, 1, InStr(DSN, ";") - 1)
pcaParseDSNString = DSN
End Function
Function pcaParseFullPath(FullPath As String, Drive As String, path As String, FileName As String, Ext As String)
Dim ColonPOs As Integer
Dim FullFileName As String
Dim DotPos As Integer
Dim LastPathPos As Integer
ColonPOs = InStr(FullPath, ":")
Drive = Mid(FullPath, 1, ColonPOs)
FullFileName = pcaStrExtract(FullPath, "\", pcaChrCount("\", FullPath) + 1)
DotPos = InStr(FullFileName, ".")
If DotPos = InStr(FullFileName, ".") Then
    Ext = Mid(FullFileName, DotPos)
    FileName = Mid(FullFileName, 1, DotPos - 1)
Else
    Ext = ""
    FileName = FullFileName
End If
'Stop
LastPathPos = Len(FullPath) - Len(FullFileName)
path = Mid(FullPath, ColonPOs + 1, LastPathPos - ColonPOs)

End Function

Function pcaProper(X)

    '  Capitalize first letter of every word in a field.
    '  Can use in an event procedure in AfterUpdate of control;
    '  for example, [Last Name] = Proper([Last Name]).
    '  Names such as O'Brien and Wilson-Smythe are properly capitalized,
    '  but MacDonald is changed to Macdonald, and van Buren to Van Buren.
    '  Note: For this function to work correctly, you must specify
    '  Option Compare Database in the Declarations section of this module.

    Dim temp$, c$, OldC$, i As Integer
    If IsNull(X) Then
        Exit Function
    Else
        temp$ = CStr(LCase(X))
        '  Initialize OldC$ to a single space because first
        '  letter needs to be capitalized but has no preceding letter.
        OldC$ = " "
        For i = 1 To Len(temp$)
            c$ = Mid$(temp$, i, 1)
            If c$ >= "a" And c$ <= "z" And (OldC$ < "a" Or OldC$ > "z") Then
                Mid$(temp$, i, 1) = UCase$(c$)
            End If
            OldC$ = c$
        Next i
        pcaProper = temp$
    End If
End Function

Function pcaConvertEmpty(Optional v, Optional DefaultVal)
On Error Resume Next
If IsMissing(v) Then v = Null
If IsMissing(DefaultVal) Then DefaultVal = Null
If pcaempty(v) Then
    pcaConvertEmpty = DefaultVal
Else
    pcaConvertEmpty = v
End If
Exit Function
End Function
Function pcaConvertNulls(Optional v, Optional DefaultVal)
If IsMissing(v) Then v = Null
If IsMissing(DefaultVal) Then DefaultVal = Null
On Error Resume Next
If IsNull(v) Then
    pcaConvertNulls = DefaultVal
Else
    pcaConvertNulls = v
End If
Exit Function
End Function
Public Function pcaFormatSSN(vSSN)
Dim vSSNOriginal As String, retval As String
If pcaempty(vSSN) Then
    pcaFormatSSN = vSSN
    Exit Function
End If
vSSNOriginal = vSSN
vSSN = Trim(pcaReplaceStr(vSSN, " ", ""))
Select Case True
Case Len(vSSN) = 9 And vSSN = Format(vSSN, "000000000")
    retval = Format(vSSN, "000-00-0000")
Case Len(vSSN) = 11 And vSSN = Format(vSSN, "000000000")
    retval = Format(vSSN, "000-00-0000")
Case Else
    retval = vSSNOriginal
End Select
pcaFormatSSN = retval
End Function
Public Function PcaFormatName(sName)
Dim retval As String
Dim sPart As String
retval = ""
Dim i As Integer
    i = 1
    sPart = pcaStrExtract(sName, " ", i)
    Do While Not pcaempty(sPart)
        sPart = LCase(sPart)
        retval = retval + UCase(Left(sPart, 1)) & Mid(sPart, 2, Len(sPart) - 1) + " "
        i = i + 1
        sPart = pcaStrExtract(sName, " ", i)
    Loop
'If Not pcaempty(sName) Then
'    sName = LCase(sName)
'    retval = UCase(Left(sName, 1)) & Mid(sName, 2, Len(sName) - 1)
'End If
    PcaFormatName = retval
End Function
Public Function PcaFormatPhoneNumber(sPhone, Optional FormatOption As Integer)
Dim retval As String, foo
Dim InitVal As String, mPhone
' Riccardo -- This function really needs to be more generic to allow for future phone
' formats.  I've added the FormatOption as an Optional parameter.  0 is the default.
' The idea is that we may want it formated as (123) 456-7890 or 123-456-7890, or...
' Also, any phone len > 10,(or > 11 if first digit is 1), where len is defined as len before any x, can't be formated.
' For example 11234567890  can't be formated because it may be a foriegn number.

If IsNull(sPhone) Then
    PcaFormatPhoneNumber = ""
    Exit Function
ElseIf pcaempty(sPhone) Then
    PcaFormatPhoneNumber = ""
    Exit Function
End If
InitVal = sPhone
mPhone = Trim(sPhone)
mPhone = pcaReplaceStr(mPhone, " ", "")
mPhone = pcaReplaceStr(mPhone, "-", "")
mPhone = pcaReplaceStr(mPhone, "(", "")
mPhone = pcaReplaceStr(mPhone, ")", "")
If InStr(mPhone, "ext") > 0 Or InStr(mPhone, "ext.") > 0 Then
    mPhone = pcaReplaceStr(mPhone, "ext", "x")
End If
Select Case True
Case pcaempty(mPhone)
    retval = ""
Case Len(mPhone) = 3
    retval = "(" & mPhone & ") "
Case Len(mPhone) = 10                                     'example "1234567896"
    retval = pcaFormatPhone10(mPhone, FormatOption)
Case Len(mPhone) = 11 And Left(mPhone, 1) = CStr(1) 'And InStr(mPhone, "-") = 0
    mPhone = Mid(mPhone, 2)
    retval = "1-(" & pcaFormatPhone10(mPhone, FormatOption)
Case Len(mPhone) > 11 And Left(mPhone, 1) = CStr(1) 'And InStr(mPhone, "-") = 0
    mPhone = pcaReplaceStr(mPhone, "x", "")                  'example "17984561234 x456"
    retval = "1-" & pcaFormatPhone10(Mid(mPhone, 2, 10), FormatOption) & " x" & Trim(pcaReplaceStr(Mid(mPhone, 12), "x", ""))
Case Len(mPhone) > 11 And Left(mPhone, 1) <> CStr(1) And InStr(mPhone, "x") > 0
    mPhone = pcaReplaceStr(mPhone, "x", "")                  'example "7984561234 x456"
    retval = pcaFormatPhone10(Mid(mPhone, 1, 10), FormatOption) & " x" & Trim(pcaReplaceStr(Mid(mPhone, 11), "x", ""))
Case Len(mPhone) = 7                                      'example "4567894"
    retval = "(   ) " & Mid(mPhone, 1, 3) & "-" & Mid(mPhone, 4, 4)
'Case Len(mPhone) = 12 And InStr(mPhone, "-") = 4           'example "781-321-5689"
'    RetVal = "(" & Left(mPhone, 3) & ") " & Mid$(mPhone, 5, 3) & "-" & Mid$(mPhone, 9, 4)
'Case Left(mPhone, 1) = "(" And Mid(mPhone, 5, 2) = ") " And Mid(mPhone, 10, 1) = "-"
'    RetVal = "(" & Mid(mPhone, 2, 3) & ") " & Mid$(mPhone, 7, 3) & "-" & Mid$(mPhone, 11, 4) 'example "(781) 321-5689"
Case Else
    'foo = MsgBox("Not recognized format!", vbOKOnly + vbExclamation, "Wrong format")
    retval = InitVal
End Select
PcaFormatPhoneNumber = retval
End Function
Private Function pcaFormatPhone10(mPhone, FormatOption As Integer)
Select Case FormatOption
Case 0
    pcaFormatPhone10 = "(" & Left$(mPhone, 3) & ") " & Mid$(mPhone, 4, 3) & "-" & Mid$(mPhone, 7, 4)
Case 1
    pcaFormatPhone10 = Left$(mPhone, 3) & "-" & Mid$(mPhone, 4, 3) & "-" & Mid$(mPhone, 7, 4)
End Select
End Function
Public Function pcaExpandPhone(mPhone As Variant) As String
pcaExpandPhone = PcaFormatPhoneNumber(mPhone)
'If pcaEmpty(mPhone) Then
'    pcaExpandPhone = ""
'    Exit Function
'End If
'pcaExpandPhone = "(" & Left$(mPhone, 3) & ") " & Mid$(mPhone, 4, 3) & "-" & Mid$(mPhone, 7, 4)
End Function
Function pcaGetToken(ByVal strValue As String, ByVal strDelimiter As String, ByVal intPiece As Integer) As Variant
Dim foo
' Given the string in strValue, and the delimiter in
' strDelimiter, find the intPiece'th token
' in the string.

' For example:
'   pcaGetToken("This is a test", " ", 4)
' will return "test".

Dim intPos As Integer
Dim intLastPos As Integer
Dim intNewPos As Integer

On Error GoTo pcaGetTokenExit

' Make sure the delimiter is just one character.
strDelimiter = Left(strDelimiter, 1)

' If the delimiter doesn't occur at all, or if
' the user's asked for a negative item, just return the

If (InStr(strValue, strDelimiter) = 0) Or (intPiece <= 0) Then
   pcaGetToken = strValue
Else
   intPos = 0
   intLastPos = 0
   Do While intPiece > 0
      intLastPos = intPos
      intNewPos = InStr(intPos + 1, strValue, strDelimiter)
      If intNewPos > 0 Then
            intPos = intNewPos
            intPiece = intPiece - 1
      Else  ' Catch the last piece, where there's no trailing token.
            intPos = Len(strValue) + 1
            Exit Do
      End If
   Loop
   If intPiece > 1 Then
      pcaGetToken = Null
   Else
      pcaGetToken = Mid$(strValue, intLastPos + 1, intPos - intLastPos - 1)
   End If
End If

pcaGetTokenExit:
   Exit Function

pcaGetTokenErr:
Select Case Err
End Select
foo = pcaStdErrMsg(Err, Error)
Resume pcaGetTokenExit
End Function


Function pcaSoundEx(vstr As String) As String
Dim rv As String
Dim s As String
Dim c As String
Dim i As Long
Dim foo
'**********************************************************************************
'* This should and can be optimized so that output is built by going thru one loop.
'**********************************************************************************
rv = vstr
'//////////////////////////////////////////////////////////
'(1)    First delete all spaces and non-alphabetic symbols.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
s = ""
For i = 1 To Len(rv)
    c = Mid(rv, i, 1)
    Select Case c
    Case "A" To "Z"
        s = s & c
    Case Else
    End Select
Next i
rv = s
'//////////////////////////////////////////////////////////
'(2)    Delete all "H" and "W", unless one is an inital letter.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
s = ""
s = Mid(rv, 1, 1)
For i = 2 To Len(rv)
    c = Mid(rv, i, 1)
    Select Case c
    Case "H", "W"
    Case Else
        s = s & c
    End Select
Next i
rv = s
'//////////////////////////////////////////////////////////
'(3) The first letter will be the first character of the
'   soundex, the remaining letters will be recoded.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
s = ""
s = Mid(rv, 1, 1)
For i = 2 To Len(rv)
    s = s & soundExRecode_(Mid(rv, i, 1))
Next i
rv = s
'//////////////////////////////////////////////////////////
'(4)    Combine all double numbers
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
s = ""
s = Mid(rv, 1, 1)
For i = 2 To Len(rv)
    If Mid(rv, i, 1) <> Mid(rv, i - 1, 1) Then
        s = s & Mid(rv, i, 1)
    End If
Next i
rv = s
'//////////////////////////////////////////////////////////
'(5)    If the first number(2nd character) is the same as
'       the code for the initial letter, delete that number
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If (Mid(rv, 2, 1) = soundExRecode_(Mid(rv, 1, 1))) Then
    rv = Mid(rv, 1, 1) & Mid(rv, 3)
End If
'//////////////////////////////////////////////////////////
'(6)    Delete the zeroes
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
rv = pcaReplaceStr(rv, "0", "")
'//////////////////////////////////////////////////////////
'(7)    Retain one letter and first three numbers,
'       concatenate zeroes on the end if needed.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
rv = Left(rv & "0000", 4)
'-----------------------------------------------------------------
EXIT_HANDLER:
pcaSoundEx = rv
Exit Function
'-----------------------------------------------------------------
ERR_HANDLER:
foo = pcaStdErrMsg(Err, Error)
Resume EXIT_HANDLER:
Resume
End Function

Private Function soundExRecode_(c As String) As String
Select Case c
Case "A", "E", "I", "O", "U", "Y": soundExRecode_ = "0"
Case "B", "F", "P", "V": soundExRecode_ = "1"
Case "C", "G", "J", "K", "Q", "S", "X", "Z": soundExRecode_ = "2"
Case "D", "T": soundExRecode_ = "3"
Case "L": soundExRecode_ = "4"
Case "M", "N": soundExRecode_ = "5"
Case "R": soundExRecode_ = "6"
Case Else: soundExRecode_ = ""
End Select
End Function
Function pcaAddQuotesToEachItemOfList(s As String, delim As String) As String
Dim retval As String, i As Integer, n As Integer
retval = ""
n = pcaChrCount(delim, s)
For i = 1 To n
    retval = retval & "'" & pcaStrExtract(s, delim, i) & "'" & delim
Next
pcaAddQuotesToEachItemOfList = retval
End Function
Function pcarTrimWhiteSpace(ByVal s As Variant) As String
On Error Resume Next
If IsNull(s) Then Exit Function
Do While True
    Select Case Asc(Right(s, 1))
    Case 10, 13, 32, 9, 0
        s = RTrim$(Mid$(s, 1, Len(s) - 1))
    Case Else
        Exit Do
    End Select
    If Len(s) = 0 Then Exit Do
Loop
pcarTrimWhiteSpace = RTrim(s)
End Function
Function pcaStrZero(n As Long, i As Integer) As String
pcaStrZero = Right("00000000000000000000" & Trim(Str(n)), i)
End Function
Function pcaUntrim(s As String, i As Integer) As String
pcaUntrim = Left$(s & Space$(i), i)
End Function
Function pcaToStr(v As Variant) As String
If IsNull(v) Then
   pcaToStr = ""
Else
   pcaToStr = CStr(v)
End If
End Function
Public Function pcaGetSQLSection(sSource, sSection As String) As String
Dim iW  As Integer, iG As Integer, iH As Integer, iO As Integer
Dim sRetVal As String
If sSection = "GroupBy" Then sSection = "Group By"
If sSection = "OrderBy" Then sSection = "Order By"

iW = InStr(sSource, "WHERE")
iG = InStr(sSource, "GROUP BY")
iH = InStr(sSource, "HAVING")
iO = InStr(sSource, "ORDER BY")
sRetVal = ""
Select Case sSection
Case "Where"
    If iW > 0 Then sRetVal = Mid(sSource, iW + Len("Where"), -iW - Len("Where") + Switch(iG > 0, iG - 1, iH > 0, iH - 1, iO > 0, iO - 1, True, Len(sSource)))
Case "Groupby", "Group By"
    If iG > 0 Then sRetVal = Mid(sSource, iG + Len("Group By"), -iG - Len("Group By") + Switch(iH > 0, iH - 1, iO > 0, iO - 1, True, Len(sSource)))
Case "Having"
    If iH > 0 Then sRetVal = Mid(sSource, iH + Len("Having"), -iH - Len("Having") + Switch(iO > 0, iO - 1, True, Len(sSource)))
Case "Orderby", "Order by"
    If iO > 0 Then sRetVal = Mid(sSource, iO + Len("Order By"), -iO - Len("Order By") + Switch(True, Len(sSource)))
Case Else:  MsgBox ("Error:Select case out of bounds. SqlSection " & sSection & " not recognized")
End Select
pcaGetSQLSection = Trim$(sRetVal)
End Function

Public Function pcaPutSQLSection(sRecordSource As String, sSQL As String, sSection As String) As String
Dim foo
'//Example1
'Dim sRefresh        As String
'Dim sSort           As String
'Dim sRecordSource   As String
'sRefresh = "CompanyID = " & CompanyID
'sSort = "CorporateData.CompanyName"
'sRecordSource = FormFilters("CorporateData", "RecordSource", Null)
'sRecordSource = PutInSource(sRecordSource, sSort, "OrderBy")
'Me.Painting = False
'Me.RecordSource = FormFilters("CorporateData", "RecordSource", sRecordSource)
'foo = FormFind(Me, sRefresh, A_FIRST)
'Me.Painting = True
'//Example2
'Dim s As String
'Dim sRecordSource As String
'If pcaempty(CurrentFilter) Then Exit Function
's = "CompanyID = " & CurrentCompanyID("")
'Select Case CurrentFilter
'Case 1
'    s = s & " and DirectorAddresses.CurrentOrFormer = " & pcaAddQuotes("Current")
'Case 2
'    s = s
'End Select
'sRecordSource = FormFilters("Directors", "RecordSource", Null)
'sRecordSource = PutInSource(sRecordSource, s, "Where")
'Me.RecordSource = FormFilters("Directors", "RecordSource", sRecordSource)
'//sSection can be "Where","Orderby","Groupby"
'//SELECT  * FROM Names WHERE (NameID = 2) GROUP BY SecuritiesID ORDER BY FullName;
Dim iW  As Integer, iG As Integer, iH As Integer, iO As Integer
Dim sSource As String, sBegining As String
Dim swhere As String, sGroupBy As String, sHaving As String, sOrderBy As String
If sSection = "GroupBy" Then sSection = "Group By"
If sSection = "OrderBy" Then sSection = "Order By"
If sSection = "OrderBy Desc" Then sSection = "Order By Desc"
sSource = pcaReplaceStr((sRecordSource), Chr(13), " ")
sSource = Trim$(pcaReplaceStr(sSource, Chr(10), " "))
sSQL = Trim$(sSQL)
If Right(sSource, 1) = ";" Then sSource = Left$(sSource, Len(sSource) - 1)
'//iW is the start of the "WHERE" component of the string, including the "WHERE".
'//sWhere is the text of the "WHERE" component of the string.
iW = InStr(sSource, "WHERE")
iG = InStr(sSource, "GROUP BY")
iH = InStr(sSource, "HAVING")
iO = InStr(sSource, "ORDER BY")
'//sBegining will equal string.start up to first occurrence of IN("WHERE","GROUB BY","HAVING","ORDER BY")
sBegining = Trim$(Left$(sSource, Switch(iW > 0, iW - 1, iG > 0, iG - 1, iH > 0, iH - 1, iO > 0, iO - 1, True, Len(sSource) + 1)))
If iW > 0 Then swhere = Trim$(Mid$(sSource, iW, -iW + Switch(iG > 0, iG - 1, iH > 0, iH - 1, iO > 0, iO - 1, True, Len(sSource) + 1)))
If iG > 0 Then sGroupBy = Trim$(Mid$(sSource, iG, -iG + Switch(iH > 0, iH - 1, iO > 0, iO - 1, True, Len(sSource) + 1)))
If iH > 0 Then sHaving = Trim$(Mid$(sSource, iH, -iH + Switch(iO > 0, iO - 1, True, Len(sSource) + 1)))
If iO > 0 Then sOrderBy = Trim$(Mid$(sSource, iO, -iO + Switch(True, Len(sSource) + 1)))
Select Case Len(sSQL)
Case 0 ' We need to replace the section with nothing...
    Select Case sSection
    Case "Where":        swhere = ""
    Case "Group By":      sGroupBy = ""
    Case "Having":       sHaving = ""
    Case "Order By Desc": sOrderBy = ""
    Case "Order By":      sOrderBy = ""
    Case Else: foo = MsgBox("Error:Select case out of bounds. SqlSection " & sSection & " not recognized")
    End Select
Case Is > 0
    Select Case sSection
    Case "Where":        swhere = " WHERE " & sSQL
    Case "Group By":      sGroupBy = " GROUP BY " & sSQL
    Case "Having":       sHaving = " HAVING " & sSQL
    Case "Order By Desc": sOrderBy = " ORDER BY " & sSQL & " DESC"
    Case "Order By":      sOrderBy = " ORDER BY " & sSQL
    Case Else: foo = MsgBox("Error:Select case out of bounds. SqlSection " & sSection & " not recognized")
    End Select
End Select
pcaPutSQLSection = Trim$(Trim$(sBegining) & " " & Trim$(swhere) & " " & Trim$(sGroupBy) & " " & Trim$(sHaving) & " " & Trim$(sOrderBy)) & ";"
End Function

Public Function pcadiv0(n, d)
'Const MB_EXCLAIM = 48
'Const ERR_DIV0 = 11, ERR_OVERFLOW = 6, ERR_ILLFUNC = 5

If IsNull(d) Then
    pcadiv0 = 0
    Exit Function
End If
    
If d = 0 Then
    pcadiv0 = 0
    Exit Function
End If
On Error Resume Next
pcadiv0 = n / d
If Err Then
    pcadiv0 = 0
End If

End Function
Function pcaMaxValue(ByVal X As Variant, ByVal Y As Variant) As Variant
If IsNull(X) Then
    pcaMaxValue = Y
    Exit Function
End If
If IsNull(Y) Then
    pcaMaxValue = X
    Exit Function
End If
If IsDate(X) And IsDate(Y) Then
    X = CVDate(X)
    Y = CVDate(Y)
End If
If X > Y Then pcaMaxValue = X Else pcaMaxValue = Y
End Function
Function pcaMinValue(ByVal X As Variant, ByVal Y As Variant) As Variant
If IsNull(X) Or IsNull(Y) Then
    pcaMinValue = Null
    Exit Function
End If
If IsDate(X) And IsDate(Y) Then
    X = CVDate(X)
    Y = CVDate(Y)
End If
If X < Y Then pcaMinValue = X Else pcaMinValue = Y
End Function
Function pcaRound(ByVal n As Variant, ByVal d As Integer) As Variant
pcaRound = Int(n * 10 ^ (d) + 0.5) / 10 ^ d
End Function
Function pcaRoundX(ByVal n As Double, nUpdown As Integer, nNearest As Double) As Double
Dim nPlus As Long
Const pcaROUNDX_UP = 1
Const pcaROUNDX_DOWN = -1
Const pcaROUNDX_NEAREST = 0

' n - Number to round
' nUpdown is one of the following (rounds up, down, or to nearest value):
    'Global Const ROUNDX_UP = 1
    'Global Const ROUNDX_DOWN = -1
    'Global Const ROUNDX_NEAREST = 0
' nNearest - the amount to round to:
'     Example: Use 1 to round to integer, .01 to round to nearest cent (hundredth), .125 to round to nearest eighth
n = n / nNearest
nPlus = Int(n + nNearest / 10000)
'If Abs(nPlus - n) < nNearest / 100 Then
If Abs(nPlus - n) < nNearest / (1000 * nNearest) Then   ' Corrected 7/22/96 by Tripp to be valid for large numbers
    pcaRoundX = nPlus * nNearest
    Exit Function
End If

Select Case nUpdown
Case pcaROUNDX_UP: n = n + 1
Case pcaROUNDX_DOWN
Case pcaROUNDX_NEAREST: n = n + 0.5
Case Else: MsgBox ("invalid data...")
End Select

pcaRoundX = Int(n) * nNearest

End Function
Function pcaRoundY(n As Long, RoundUp_RoundDown As Integer, lMultiple As Long) As Long
'EMAN.08.05.96
'//RoundUp_RoundDown enum {1,-1}

Select Case RoundUp_RoundDown
Case 1
    pcaRoundY = n + lMultiple - (n Mod lMultiple)
Case -1
    pcaRoundY = n - (n Mod lMultiple)
Case Else
    MsgBox "error in select case"
End Select
End Function
Public Function PcaShowPcaDebug() As Boolean
PcaShowPcaDebug = pcaIsFile("C:\pcaDebug")
End Function
Public Function PcaRaiseError(nErr As Integer, sProjectClass As String, sErrorDescription As String)
Dim foo
If PcaShowPcaDebug Then
    foo = pcaWarningMsg("Error:" & nErr & vbCrLf & "Project.Class:" & sProjectClass & vbCrLf & "Error.Description:" & sErrorDescription, "Pca Raise Error")
End If
Err.Raise nErr, sProjectClass, sErrorDescription
End Function
Function Pca1000DebugMsg(MSG As String)
Dim foo
If pcaIsFile("C:\pcaDebug") Then foo = MsgBox(MSG, vbInformation, "pca1000 Debug Msg, Rename file c:\pcaDebug to not show")

'If AccessUserName() = "pca1000" Then
'    If dlookup("Pca1000Debug", "z_PcaMdbTag") Then
'        foo = MsgBox(msg, vbInformation, "pca1000 Debug Msg")
'    End If
'Else
'    If dlookup("UserDebug", "z_PcaMdbTag") Then
'        foo = MsgBox(msg, vbInformation, "User Debug Msg")
'    End If
'End If
Exit Function
End Function
Public Function pcaStdErrMsg(Optional nErr As Long, Optional sError As String)
   '% requires:   nothing
   '% modifies:   nothing
   '% effects:    Displays a messagebox describing the error that has occurred.
   '              If the arguments are not provided, the message will be based on the
   '              current state of the Err object
'DoCmd.Echo True
If IsMissing(nErr) Then nErr = Err.Number
If IsMissing(sError) Then sError = Err.Description
'DoCmd.Hourglass False
Dim foo As Integer, MSG As String
'Dim DAOError As String
'Dim DAOErrorNumber As Long
'Dim DAOErrorDescription As String

MSG = "Error: " & nErr & Chr(13) & Chr(10) & sError
'If dao.Errors.Count > 0 Then
'    DAOError = dao.Errors(0)
'    DAOErrorNumber = dao.Errors(0).Number
'    DAOErrorDescription = dao.Errors(0).Description
'End If
'If DAOError <> "" Then
'    msg = msg & Chr(13) & Chr(10) & "(Last DAO Error was:" & DAOErrorNumber & " " & DAOErrorDescription & ")"
'End If
Debug.Print "StdErrMsg:" & MSG

Select Case nErr
Case 3146 ' Odbc Call Failed
    Beep
    Beep
    Beep
    foo = MsgBox(MSG, 48, "PCA ODBC Error Message...")
Case Else

    Beep
    foo = MsgBox(MSG, 48, "PCA Error Message...")
End Select

End Function
Function pcaWarningMsg(s As String, Optional sTitle As String = "PCA Warning Message:") As Integer
pcaWarningMsg = MsgBox(s, vbExclamation, sTitle)
End Function
Function ask(s As String, Optional sTitle As String = "?", Optional def As Integer = 0) As Integer
ask = (MsgBox(s, vbQuestion + vbYesNo + IIf(def, 0, vbDefaultButton2), sTitle) = vbYes)
End Function
Function AskYesNoCancel(s As String, sTitle As String, def As Integer) As Integer
'--- def = 0 for first default button
'--- def = 256 for second default button
'--- def = 512 for third default button
AskYesNoCancel = MsgBox(s, 35 + def, sTitle)
End Function
Function pcaFileSelect(Optional vDefaultFilter As String, Optional vDefaultFileName As String, _
                            Optional vDirectory As String, Optional vDialogTitle As String) As String
On Error GoTo ERR_HANDLER
Dim FileFilter As String
Dim rv As String
Dim strFilter As String
Dim lngFlags As Long
Dim lngFilterIndex As Long

    lngFlags = ahtOFN_FILEMUSTEXIST Or _
                ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
                
    strFilter = ahtAddFilterItem(strFilter, "Microsoft Access (*.mda, *.mdb, *.mde)", _
                    "*.MDA;*.MDB;*.MDE")
    strFilter = ahtAddFilterItem(strFilter, "Excel Files (*.xls)", _
                    "*.XLS")
    strFilter = ahtAddFilterItem(strFilter, "dBASE Files (*.dbf)", "*.DBF")
    strFilter = ahtAddFilterItem(strFilter, "Text Files (*.txt, *.csv, *.tab, *.asc)", _
                    "*.TXT;*.CSV;*.TAB;*ASC")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
                
                
    Select Case vDefaultFilter
    Case "MSAccess", "Access":  lngFilterIndex = 1
    Case "excel":   lngFilterIndex = 2
    Case "Text": lngFilterIndex = 4
    Case Else
        lngFilterIndex = 5
    End Select
    
    If pcaempty(vDirectory) Then vDirectory = "C:\"
    If pcaempty(vDialogTitle) Then vDialogTitle = "Please select a file"
    
    rv = ahtCommonFileOpenSave(InitialDir:=vDirectory, _
        Filter:=strFilter, FilterIndex:=lngFilterIndex, Flags:=lngFlags, _
        DialogTitle:=vDialogTitle, FileName:=vDefaultFileName, OpenFile:=True)
    
    pcaFileSelect = rv
'--------------------------------
EXIT_HANDLER:
    Exit Function
'--------------------------------
ERR_HANDLER:
    Call pcaStdErrMsg(Err.Number, Err.Description)
    Resume EXIT_HANDLER
    Resume
End Function
Function pcaFileSaveAs(Optional vDefaultFilter As String, Optional vDefaultFileName As String, _
                            Optional vDirectory As String, Optional vDialogTitle As String) As String
On Error GoTo ERR_HANDLER
Dim FileFilter As String
Dim rv As String
Dim strFilter As String
Dim lngFlags As Long
Dim lngFilterIndex As Long

    lngFlags = ahtOFN_FILEMUSTEXIST Or _
                ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
                
    strFilter = ahtAddFilterItem(strFilter, "Microsoft Access (*.mda, *.mdb, *.mde)", _
                    "*.MDA;*.MDB;*.MDE")
    strFilter = ahtAddFilterItem(strFilter, "Excel Files (*.xls)", _
                    "*.XLS")
    strFilter = ahtAddFilterItem(strFilter, "dBASE Files (*.dbf)", "*.DBF")
    strFilter = ahtAddFilterItem(strFilter, "Text Files (*.txt, *.csv, *.tab, *.asc)", _
                    "*.TXT;*.CSV;*.TAB;*ASC")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
                
                
    Select Case vDefaultFilter
    Case "MSAccess", "Access":  lngFilterIndex = 1
    Case "excel":   lngFilterIndex = 2
    Case "Text": lngFilterIndex = 4
    Case Else
        lngFilterIndex = 5
    End Select
    
    If pcaempty(vDirectory) Then vDirectory = "C:\"
    If pcaempty(vDialogTitle) Then vDialogTitle = "Please save a file as"
    
    rv = ahtCommonFileOpenSave(InitialDir:=vDirectory, _
        Filter:=strFilter, FilterIndex:=lngFilterIndex, Flags:=lngFlags, _
        DialogTitle:=vDialogTitle, FileName:=vDefaultFileName, OpenFile:=False)
    
    pcaFileSaveAs = rv
'--------------------------------
EXIT_HANDLER:
    Exit Function
'--------------------------------
ERR_HANDLER:
    Call pcaStdErrMsg(Err.Number, Err.Description)
    Resume EXIT_HANDLER
    Resume
End Function


Function pcaIsFile(sFile As Variant, Optional DispMsg As Boolean = False) As Integer
On Error GoTo ERR_HANDLER
Dim MSG As String, foo
Const ERR_DISKNOTREADY = 71, ERR_DEVICEUNAVAILABLE = 68
Const MB_EXCLAIM = 48, MB_STOP = 16, MB_OK_CANCEL = 1, BUTTON_OK = 1
pcaIsFile = (GetAttr(sFile) And vbDirectory) = 0
'=========================================================
EXIT_FUNCTION:
Exit Function
'=========================================================
ERR_HANDLER:
Select Case Err
'Case ERR_DISKNOTREADY
'   Msg = "Put a floppy disk in the drive and close the drive door."
'   If MsgBox(Msg, MB_EXCLAIM + MB_OK_CANCEL) = BUTTON_OK Then
'      Resume
'   Else
'      Resume Next
'   End If
Case 53
    pcaIsFile = False
    Resume EXIT_FUNCTION
Case ERR_DEVICEUNAVAILABLE
    pcaIsFile = False
    If DispMsg Then
        MSG = "This drive or path does not exist: " + sFile
        MsgBox MSG, MB_EXCLAIM
        Resume Next
    End If
    Resume EXIT_FUNCTION
Case 76           'Path not found
    pcaIsFile = False
    If DispMsg Then
        MSG = "This path does not exist: " + sFile
        MsgBox MSG, MB_EXCLAIM
        Resume Next
    End If
    Resume EXIT_FUNCTION
Case Else
    If DispMsg Then
        foo = pcaStdErrMsg(Err, Error)
    End If
    pcaIsFile = False
    Resume EXIT_FUNCTION
End Select
foo = pcaStdErrMsg(Err, Error)
Resume EXIT_FUNCTION:
Resume
End Function
Function pcaDeleteFile(fName As String) As Integer
On Error GoTo ERR_HANDLER
If fName = "" Then
    pcaDeleteFile = False
    Exit Function
End If
If Not pcaIsFile(fName) Then
    pcaDeleteFile = True
    Exit Function
End If
Kill fName
'--------------------------------------------
EXIT_FUNCTION:
pcaDeleteFile = Not pcaIsFile(fName)
Exit Function
'--------------------------------------------
ERR_HANDLER:
Select Case Err
Case 75  'Path/File access error
    pcaDeleteFile = False
    Resume EXIT_FUNCTION
End Select
Call pcaStdErrMsg(Err, Error)
Resume EXIT_FUNCTION
Resume
End Function
Function pcaGetTempFileName(ByVal DriveLetter As Integer, ByVal Prefix As String, ByVal Unique As Integer) As String
'Dim buffer As String * 255, ReturnValue As Integer
'ReturnValue = GetTempFileName(DriveLetter, Prefix, Unique, buffer)
'fgetTempFileName = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
End Function
Function pcaFileOpen(sFileName As String, OpenMode As Integer) As Integer
Dim fh As Integer, foo
On Error GoTo ERR_HANDLER:
Const FOPEN_WRITE = 1
Const FOPEN_READWRITE = 2
Const FOPEN_READ = 0
fh = FreeFile
Select Case OpenMode
Case FOPEN_WRITE
    Open sFileName For Output As fh
Case FOPEN_READWRITE
    Open sFileName For Output As fh
Case Else ' FOPEN_READ
    Open sFileName For Input Access Read As fh
End Select
EXIT_FUNCTION:
pcaFileOpen = fh
Exit Function
'----------------------------------------------------------------
ERR_HANDLER:
fh = -1
Select Case Err
End Select
Call pcaStdErrMsg(Err, Error)
Resume EXIT_FUNCTION
Resume
End Function
Function pcaFileRename(fNameOld As String, fNameNew As String, Optional DispMsgs As Boolean) As Integer
Dim foo
On Error Resume Next
If pcaempty(fNameOld) Or pcaempty(fNameNew) Then
    MsgBox ("Must supply Old and New filename...")
    pcaFileRename = False
    Exit Function
End If

If Not pcaIsFile(fNameOld) Then
    MsgBox ("File " & fNameOld & " does not exist, cannot be renamed...")
    pcaFileRename = False
    Exit Function
End If
If pcaIsFile(fNameNew) Then
    MsgBox ("File " & fNameNew & " does already exists, cannot rename " & fNameOld & "...")
    pcaFileRename = False
    Exit Function
End If

Name fNameOld As fNameNew

If Err Then
    Call pcaStdErrMsg(Err, Error)
    pcaFileRename = False
    Exit Function
End If

pcaFileRename = Not pcaIsFile(fNameOld) And pcaIsFile(fNameNew)
Exit Function

End Function

Function pcaGetFilesForPattern(fPattern As String) As String
Dim foo
On Error GoTo ERR_GetFilesForPattern
Dim sFile As String
Dim sAllFiles As String
sFile = Dir$(fPattern)
sAllFiles = ""
Do While Not pcaempty(sFile)
    sAllFiles = sFile & ";" & sAllFiles
    sFile = Dir$
Loop
pcaGetFilesForPattern = sAllFiles
Exit Function
'----------------------------------
ERR_GetFilesForPattern:
Select Case Err
Case 76     'Path not found
    sFile = ""
    Resume Next
End Select
Call pcaStdErrMsg(Err, Error)
Resume Next
Resume
End Function

Function pcaIsDir(sPath As Variant) As Integer
Dim foo
On Error GoTo ERR_Pca_IsDir
Dim retval As Integer
If pcaempty(sPath) Then
    retval = False
    GoTo EXIT_Pca_IsDir
Else
    retval = True
End If
retval = Not pcaempty(Dir(sPath, vbDirectory))
'ChDir sPath
EXIT_Pca_IsDir:
pcaIsDir = retval
Exit Function
'-----------------------------------
ERR_Pca_IsDir:
Select Case Err
Case 76     'Path not found
    retval = False
    Resume Next
End Select
Call pcaStdErrMsg(Err, Error)
Resume EXIT_Pca_IsDir
Resume
End Function

Function pcaSetCurDir(sPath As String) As Integer
ChDir sPath
If CurDir = sPath Then
    pcaSetCurDir = True
Else
    pcaSetCurDir = False
End If

End Function

Function pcaReadFileIntoString(strFile As String) As String
Dim fh As Long, foo
Dim buffer As String, retval As String
On Error GoTo ERR_HANDLER:
fh = pcaFileOpen(strFile, 0)
If fh < 0 Then
    MsgBox "Couldn't open file: " & strFile
    GoTo EXIT_FUNCTION
End If
Seek fh, 1
Do While Not EOF(fh)
    Line Input #fh, buffer
    retval = retval & buffer & vbCrLf
Loop
EXIT_FUNCTION:
If fh <> -1 Then Close fh

pcaReadFileIntoString = retval
Exit Function

ERR_HANDLER:
Call pcaStdErrMsg(Err, Error)
Resume EXIT_FUNCTION
End Function
Public Function pcaObjIsNothing(O As Object) As Boolean
Dim foo
If O Is Nothing Then
    pcaObjIsNothing = True
    Exit Function
End If
On Error Resume Next
foo = O.Name
pcaObjIsNothing = (Err <> 0)
Err = 0
Exit Function
End Function
'Function pcaOpenwsJetApp() As Integer
'Dim foo
'On Error GoTo ERR_HANDLER
'Dim sDataFile  As String
'Dim e As String
'If Not pcaWsOpen("wsJetApp") Then
'    foo = MsgBox("Invalid Account Name or Password...", "Access Denied!")
'    GoTo Exit_Function:
'End If
'pcaOpenwsJetApp = True
'Exit_Function:
'Exit Function
''------------------------------------------------------------
'ERR_HANDLER:
'Select Case Err
'Case 3029   '  Not a valid account name or password.
'    MsgBox (Err & Error)
'    Resume Exit_Function
'End Select
'
'foo = pcaStdErrMsg(Err, Error)
'
'Resume Exit_Function
'Resume
'
'End Function
'Function pcaOpenwsJetData() As Boolean
'Dim foo
'On Error GoTo ERR_HANDLER
'Dim sDataFile  As String
'Dim e As String
'If Not pcaWsOpen("wsJetData") Then
'    foo = MsgBox("Invalid Account Name or Password...", "Access Denied!")
'    GoTo Exit_Function:
'End If
'pcaOpenwsJetData = True
'
'Exit_Function:
'Exit Function
''------------------------------------------------------------
'ERR_HANDLER:
'Select Case Err
'Case 3029   '  Not a valid account name or password.
'    MsgBox (Err & Error)
'    Resume Exit_Function
'End Select
'
'foo = pcaStdErrMsg(Err, Error)
'
'Resume Exit_Function
'Resume
'
'End Function
Public Function pcalibINIGetSection(pSection As String, pFile As String) As String
'// Returns a string containing every key=value pair. Each pair is terminated by a \0. (Asc = 0).
'// Max length = 260
Dim sReturn As String * MAX_PATH
Dim sLen As Long
sLen = GetPrivateProfileSection(pSection, sReturn, MAX_PATH, pFile)
pcalibINIGetSection = Mid(sReturn, 1, sLen)
End Function
Public Function pcalibINIGetString(pSection As String, pKey As String, pFile As String, Optional pDefault As String = "") As String
Dim sReturn As String * MAX_PATH
Dim sLen As Long
sLen = GetPrivateProfileString(pSection, pKey, pDefault, sReturn, MAX_PATH, pFile)
pcalibINIGetString = Mid(sReturn, 1, sLen)
End Function
Public Function pcalibINISetString(pSection As String, pKey As String, pValue As String, pFile As String)
Dim sLen As Long
sLen = WritePrivateProfileString(pSection, pKey, pValue, pFile)
End Function
Public Function pcalibINISetSection()
Err.Raise 0, "pcalibPCA_INIFileAPI.pcalibINISetSection()", "You tried to use the function 'pcalibINISetSection' which is not implemented yet."
End Function
Public Function pcaFormatZipPlus4(Optional sZip As Variant) As String
Dim tempZip As String
tempZip = ""

On Error GoTo Err_pcaFormatZipPLus4
    
Select Case True
Case pcaempty(sZip)
    tempZip = "" 'sZip
Case Len(sZip) <= 5
    tempZip = sZip
Case Len(sZip) = 9
    tempZip = Left(sZip, 5)
    tempZip = tempZip & "-" & Right(sZip, 4)
Case Else
    tempZip = sZip
End Select

pcaFormatZipPlus4 = tempZip

EXIT_FUNCTION:
Exit Function
'----------------
Err_pcaFormatZipPLus4:
foo = pcaStdErrMsg(Err, Error)
Resume Next
Resume
End Function
Function pcaFormatYear(Optional sYear As Variant) As String
Dim tempYear As String
On Error GoTo Err_pcaFormatYear

If pcaempty(sYear) Then
    Exit Function
End If

If Len(sYear) = 1 Then
    sYear = "0" & sYear
End If

tempYear = sYear

Select Case True
Case Len(sYear) = 2
    If CInt(sYear) < 20 Then
        tempYear = "20" & sYear
    Else
        tempYear = "19" & sYear
    End If
Case Else
End Select

pcaFormatYear = tempYear

EXIT_FUNCTION:
Exit Function
'----------------
Err_pcaFormatYear:
foo = pcaStdErrMsg(Err, Error)
Resume Next
Resume
End Function
Function pcaStringFromGUID(sGUID As Variant) As String
Dim rv As String
On Error GoTo Err_pcaStringFromGUID

rv = StringFromGUID(sGUID)
rv = pcaReplaceStr(rv, "{", "")
rv = pcaReplaceStr(rv, "}", "")
rv = pcaReplaceStr(rv, "guid", "")
rv = pcaReplaceStr(rv, " ", "")

EXIT_FUNCTION:
pcaStringFromGUID = rv
Exit Function
'--------------
Err_pcaStringFromGUID:
'foo = stdErrMsg(Err, Error)
rv = ""
Resume Next
Resume
End Function
Function PcaFormatCreditCard(sCreditCard As String) As String
On Error GoTo ERR_HANDLER
Dim rv As String
Dim i As Integer
    rv = sCreditCard
    rv = pcaReplaceStr(rv, "-", "")
    rv = pcaReplaceStr(rv, " ", "")
    rv = pcaReplaceStr(rv, ".", "")
    rv = pcaReplaceStr(rv, "/", "")
    rv = Left(rv, 16)
    rv = Left(rv, 4) & "-" & Mid(rv, 5, 4) & "-" & Mid(rv, 9, 4) & "-" & Mid(rv, 13)
EXIT_HANDLER:
    PcaFormatCreditCard = rv
    Exit Function
ERR_HANDLER:
    rv = ""
    foo = pcaStdErrMsg(Err, Error)
    Resume EXIT_HANDLER
    Resume
End Function

Public Function ahtCommonFileOpenSave( _
            Optional ByRef Flags As Variant, _
            Optional ByVal InitialDir As Variant, _
            Optional ByVal Filter As Variant, _
            Optional ByVal FilterIndex As Variant, _
            Optional ByVal DefaultExt As Variant, _
            Optional ByVal FileName As Variant, _
            Optional ByVal DialogTitle As Variant, _
            Optional ByVal hwnd As Variant, _
            Optional ByVal OpenFile As Variant) As Variant
' This is the entry point you'll use to call the common
' file open/save dialog. The parameters are listed
' below, and all are optional.
'
' In:
' Flags: one or more of the ahtOFN_* constants, OR'd together.
' InitialDir: the directory in which to first look
' Filter: a set of file filters, set up by calling
' AddFilterItem. See examples.
' FilterIndex: 1-based integer indicating which filter
' set to use, by default (1 if unspecified)
' DefaultExt: Extension to use if the user doesn't enter one.
' Only useful on file saves.
' FileName: Default value for the file name text box.
' DialogTitle: Title for the dialog.
' hWnd: parent window handle
' OpenFile: Boolean(True=Open File/False=Save As)
' Out:
' Return Value: Either Null or the selected filename
On Error GoTo ERR_HANDLER
Dim OFN As tagOPENFILENAME
Dim strFileName As String
Dim strFileTitle As String
Dim fResult As Boolean
    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(FileName) Then FileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(hwnd) Then hwnd = Application.hWndAccessApp
    If IsMissing(OpenFile) Then OpenFile = True
    ' Allocate string space for the returned strings.
    strFileName = Left(FileName & String(256, 0), 256)
    strFileTitle = String(256, 0)
    ' Set up the data structure before you call the function
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hwnd
        .strFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = strFileName
        .nMaxFile = Len(strFileName)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir
        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        '.strCustomFilter = ""
        '.nMaxCustFilter = 0
        .lpfnHook = 0
        'New for NT 4.0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
    ' This will pass the desired data structure to the
    ' Windows API, which will in turn it uses to display
    ' the Open/Save As Dialog.
    If OpenFile Then
        fResult = aht_apiGetOpenFileName(OFN)
    Else
        fResult = aht_apiGetSaveFileName(OFN)
    End If

    ' The function call filled in the strFileTitle member
    ' of the structure. You'll have to write special code
    ' to retrieve that if you're interested.
    If fResult Then
        ' You might care to check the Flags member of the
        ' structure to get information about the chosen file.
        ' In this example, if you bothered to pass in a
        ' value for Flags, we'll fill it in with the outgoing
        ' Flags value.
        If Not IsMissing(Flags) Then Flags = OFN.Flags
        ahtCommonFileOpenSave = TrimNull(OFN.strFile)
    Else
        ahtCommonFileOpenSave = vbNullString
    End If
EXIT_HANDLER:
    Exit Function
ERR_HANDLER:
    'fail to open the common file open save dialog
    ahtCommonFileOpenSave = InputBox("Please Enter File Name:", "Filename?", FileName)
    Resume EXIT_HANDLER
    Resume
End Function

Public Function ahtAddFilterItem(strFilter As String, _
    strDescription As String, Optional varItem As Variant) As String
' Tack a new chunk onto the file filter.
' That is, take the old value, stick onto it the description,
' (like "Databases"), a null character, the skeleton
' (like "*.mdb;*.mda") and a final null character.

    If IsMissing(varItem) Then varItem = "*.*"
    ahtAddFilterItem = strFilter & _
                strDescription & vbNullChar & _
                varItem & vbNullChar
End Function

Public Function TrimNull(ByVal strItem As String) As String
Dim intPos As Integer
    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        TrimNull = Left(strItem, intPos - 1)
    Else
        TrimNull = strItem
    End If
End Function

Function PcaRunQuery(s As String)
On Error GoTo ERR_HANDLER
DoEvents
Debug.Print "Running: " & s
Echo True, "Running: " & s
'DoCmd.OpenQuery s
CurrentDb.Execute s
DoEvents
Exit Function

SELECTQUERY:
DoCmd.OpenQuery s
Exit Function

ERR_HANDLER:
If Err = 3065 Then
    Err = 0
    Resume SELECTQUERY
End If
MsgBox (Err & "-" & Error)
Resume Next
Resume

End Function

Public Function PcaDetachAllLinkTables()
On Error GoTo ERR_HANDLER
Dim rs As DAO.Recordset
Dim sql As String
Dim db As Database
Dim rv As Boolean


    rv = False
    Set db = CurrentDb()
    
    sql = ""
    sql = sql & "select ConnectAs from z_PCADataSources_TableList"
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        sql = ""
        sql = sql & "DROP TABLE [" & rs("ConnectAs") & "]"
        
        db.Execute sql
        rs.MoveNext
    Loop
    rv = True
EXIT_HANDLER:
    PcaDetachAllLinkTables = rv
    Exit Function
ERR_HANDLER:
    Resume Next
    rv = False
    foo = pcaStdErrMsg(Err, Error)
    GoTo EXIT_HANDLER
End Function

Public Function PcaAttachLinkTables() As Boolean
On Error GoTo ERR_HANDLER
' we will need to create this table using DAO
Dim tdf As DAO.TableDef

' Some variable to make the code more generic
Dim strConnectionString As String
Dim strNameInAccess As String
Dim strNameInSQLServer As String
Dim strKey As String

Dim rs As DAO.Recordset
Dim sql As String
Dim rv As Boolean


    rv = False

    ' specify the tables you want to link. The table can be
    ' known by a different name in Access than the name in SQL server

    
    sql = ""
    sql = sql & "SELECT * FROM z_PCADataSources_TableList"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.EOF Then rs.MoveFirst
    
    Do While Not rs.EOF
        strNameInAccess = rs("ConnectAs")
        strNameInSQLServer = rs("ForeignTableName")
        strKey = Nz(rs("UniqueID"), "")
        

        ' Create a table using DAO give it a name in Access. Connect it to the SQL Server database. Say which table it links to in SQL Server.

        Set tdf = CurrentDb.CreateTableDef(strNameInAccess)

        tdf.Connect = "ODBC;DRIVER={SQL Server};SERVER=" & PcaGetSQLServerName() & ";DATABASE=TateBywater;Trusted Connection=No;Uid=TateBywaterSQLUser;Pwd=tatebywater1234"
        

        tdf.SourceTableName = strNameInSQLServer

        ' Add this table Definition to the collection of Access tables

        CurrentDb.TableDefs.Append tdf
        'DoCmd.TransferDatabase acLink, "ODBC Database", strConnectionString, acTable, strNameInSQLServer, strNameInAccess
        
        
        ' Now create a unique key for this table by running this SQL

        If strKey <> "" Then
            DoCmd.RunSQL "CREATE UNIQUE INDEX UniqueIndex ON " & strNameInAccess & " (" & strKey & ")"
        End If

        rs.MoveNext
    Loop
    rv = True
EXIT_HANDLER:
    PcaAttachLinkTables = rv
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    Select Case Err
    Case 3010 'table already exists
        Resume Next
    Case 3283 'table already exists
        Resume Next
    Case Else
        foo = pcaStdErrMsg(Err, Error)
        rv = False
    End Select
    GoTo EXIT_HANDLER
Resume
End Function

Public Function dbINIGetSetting(pSection As String, _
                     pKey As String, _
            Optional pDefault As String = "", _
            Optional pAutoAdd As Boolean = False _
) As String
On Error GoTo ERR_HANDLER
Dim db As Database
Dim sql As String
Dim rs As DAO.Recordset

    Set db = CurrentDb()
    sql = ""
    sql = sql & "SELECT * FROM z_PCASettings"
    

    Set rs = db.OpenRecordset(sql)

    If Not rs.EOF Then rs.MoveFirst
    Do Until rs.EOF
        If rs("INISection") = pSection And rs("INIKey") = pKey Then
            dbINIGetSetting = rs!INIDescription
            Exit Function
        End If
        rs.MoveNext
    Loop
    dbINIGetSetting = "" & pDefault
EXIT_HANDLER:
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
Err.Raise vbObjectError + Err.Number, Err.Source, Err.Description
Resume
End Function

Public Function PcaGetConnnectionString() As String
On Error GoTo ERR_HANDLER
Dim db As Database
Dim sql As String
Dim rs As DAO.Recordset
Dim sConnectionString As String

Dim sApplicationStatus As String

    sApplicationStatus = dbINIGetSetting("ApplicationProperties", "ApplicationStatus")
    
    

    Set db = CurrentDb()
    sql = ""
    sql = sql & "SELECT * FROM z_PCADataSources WHERE ApplicationStatus = " & pcaAddQuotes(sApplicationStatus)
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        rs.MoveFirst
        sConnectionString = rs("DataSourceConnectString")
    Else
        sConnectionString = ""
    End If
    
    
EXIT_HANDLER:
    PcaGetConnnectionString = sConnectionString
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    foo = pcaStdErrMsg(Err, Error)
    sConnectionString = ""
    GoTo EXIT_HANDLER
Resume
End Function

Public Function PcaGetSQLServerName() As String
On Error GoTo ERR_HANDLER
Dim db As Database
Dim sql As String
Dim rs As DAO.Recordset
Dim sServerName As String

Dim sApplicationStatus As String

    sApplicationStatus = dbINIGetSetting("ApplicationProperties", "ApplicationStatus")
    
    

    Set db = CurrentDb()
    sql = ""
    sql = sql & "SELECT * FROM z_PCADataSources WHERE ApplicationStatus = " & pcaAddQuotes(sApplicationStatus)
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        rs.MoveFirst
        sServerName = rs("SQLServerName")
    Else
        sServerName = ""
    End If
    
    
EXIT_HANDLER:
    PcaGetSQLServerName = sServerName
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    foo = pcaStdErrMsg(Err, Error)
    sServerName = ""
    GoTo EXIT_HANDLER
Resume
End Function

Public Function PcaGetApplicationVersion() As String
On Error GoTo ERR_HANDLER
Dim db As Database
Dim sql As String
Dim rs As DAO.Recordset
Dim sApplicationVersion As String


    sApplicationVersion = dbINIGetSetting("ApplicationProperties", "ApplicationVersion")
    
    
    
EXIT_HANDLER:
    PcaGetApplicationVersion = sApplicationVersion
    Exit Function
Exit Function
'---------------------------------------------------------------------------------
ERR_HANDLER:
    foo = pcaStdErrMsg(Err, Error)
    sApplicationVersion = ""
    GoTo EXIT_HANDLER
Resume
End Function




'OLD VERSION of pcaFileSelect
'Function pcaFileSelect(Optional vDefaultFilter As String, Optional vDefaultFileName As String) As String
'' ### WORK IN PROGRESS ###
'   '% requires:   nothing
'   '% modifies:   nothing
'   '% effects:    Displays a GUI interface for the user to select a file for the calling application
'   '              to use. The type of display will vary depending the system on which the project
'   '              is running. The display could be as simple as an input box, or as complex as the
'   '              commonDialog control. vDefaultFilter might be used for the filtering of the files the
'   '              user will be able to select from. vDefaultFileName might be used to select the default
'   '              file/location for the display. If there is an error or the user does not select a file
'   '              the return value will be the empty string "".
'On Error GoTo ERR_HANDLER
'Dim obj As Object
'Dim FileFilter As String
'Dim rv As String
'
''// Try and make the commonDialog fileSelector since it is the best.
'On Error Resume Next
'Dim FailedToCreateCustomControl As Boolean
''Set obj = CreateObject("always_fail_so_I_can_test_MSComDlg.CommonDialog")
'
'Set obj = CreateObject("MSComDlg.CommonDialog")
'
'FailedToCreateCustomControl = (Err.Number <> 0)
'On Error GoTo ERR_HANDLER
'
'If FailedToCreateCustomControl Then
'   '// creation failed. So use or own form
''   Dim f As New frmPCAFileIOUtils_pcaFileSelect
''   f.myDefaultFileName = vDefaultFileName
''   f.myDefaultFilter = vDefaultFilter
''   Set f.myRecentFileList = pcaFileSelect_RecentFileList
''   f.Show vbModal
''   rv = f.myReturnValue
''   Set f = Nothing
''   '// Attempt to add the file to the recent file list
'   rv = InputBox("Please Enter File Name:", "Filename?", vDefaultFileName)
'   If pcaIsFile(rv) Then
'      On Error Resume Next
'      pcaFileSelect_RecentFileList.Add rv, rv
'      On Error GoTo ERR_HANDLER
'   End If
'
'Else
'   '// creation succeeded
'   Select Case vDefaultFilter
'   Case "MSAccess", "Access":   FileFilter = "Databases (*.mdb)|*.mdb|Compiled (*.mde)|*.mde|All Files (*.*)|*.*"
'   Case "Word", "WinWord":   FileFilter = "Documents (*.doc)|*.doc|Rich Text (*.rtf)|*.rtf|Text (*.txt,*.sdf,*.ini,*.asc)|*.txt;*.sdf;*.ini;*.asc|All Files (*.*)|*.*"
'   Case "excel":   FileFilter = "Excel (*.xls)|*.xls|All Files (*.*)|*.*"
'   Case Else
'       FileFilter = vDefaultFilter
'   End Select
'   'obj.Filter = s '// Sample Filter -> Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico
'   obj.DialogTitle = "Please select a file."
'   obj.InitDir = ""
'   obj.Filter = FileFilter
'   If pcaIsFile(vDefaultFileName) Then
'       obj.FileName = vDefaultFileName
'   End If
'   obj.MaxFileSize = 256
'   obj.ShowOpen
'   rv = obj.FileName
'End If
'pcaFileSelect = rv
''--------------------------------
'EXIT_HANDLER:
'Exit Function
''--------------------------------
'ERR_HANDLER:
'Call pcaStdErrMsg(Err.Number, Err.Description)
'Resume EXIT_HANDLER
'Resume
'End Function
