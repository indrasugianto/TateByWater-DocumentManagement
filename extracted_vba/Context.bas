' Component: Context
' Type: module
' Lines: 44
' ============================================================

Option Compare Database

'////////////////////////////////////////////////////////////////////////
'/ Store all the variables here which you need to maintain through out //
'/ Context module will hold all the static variables                   //
'////////////////////////////////////////////////////////////////////////

Public IsDisableEvents As Boolean

Public Enum MsgType
    Done = 1
    Saved = 2
    Canceled = 3
    NoDataQuery = 4
End Enum

Public YesNo_Value As Integer

Public message As String
Public Message1 As String

Public SelectedPath As String '' will be used to supply path from one location to another. Just a temp variable, can be used by multiple locations


Public Enum LoginModeLevel
    Strict = 1
    Light = 2
End Enum

Public ColumnHeadings() As String

Public Const Debugging As Boolean = True

Public TableNameToDeleteUpload As String '  the variable which will be used in Delete Command

Public Const CompanyName As String = "Tate Bywater CMS"

Public Property Get GetCompanyName() As String
    GetCompanyName = CompanyName
End Property



