Attribute VB_Name = "modMain"


'Registry related constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

'Files & folders lock related constants
Public Const FILE_LIST_DIRECTORY = &H1
Public Const FILE_SHARE_READ = &H1&
Public Const FILE_SHARE_DELETE = &H4&
Public Const OPEN_EXISTING = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000

Public File_Share_Flag As Long
Public CurrentUser As Userinfo
Public reg As New Registry.regedits
Public rs_user As New ADODB.Recordset
Public rs_secuser As New ADODB.Recordset
Public rs_lockinfo As New ADODB.Recordset
Public cn As New ADODB.Connection
Public Const strAppTitle = "Wildfox Files & Folders Locker Pro V1.0.0.1"
Public C_Attempt As Integer
Public Hasrun As Boolean
Public DBlocked As Boolean
Public Dbhandle As String

Public FromAddnew As Boolean
Public EditFlag As Boolean
Public AddnewFlag As Boolean

Public Type Userinfo
username As String
password As String
user_timelogin As String
usergroup As String
End Type


'Connect to database
Public Function ConnectToDB() As Boolean

On Error GoTo OpenErr

Dim DatabasePath

Set cn = New ADODB.Connection
DatabasePath = App.Path & "\" & "lockinfo.mdb"

    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & ";Persist Security Info=False;Jet OLEDB:Database Password = wildfox"
    ConnectToDB = True
Exit Function

OpenErr:

    MsgBox "Error Opening " & DatabasePath & vbNewLine & Err.Description, vbCritical, "Error Opening Database"
    ConnectToDB = False

End Function

Sub UnloadAllForms()
Dim frm As Form

For Each frm In Forms
Unload frm
Set frm = Nothing
Next

End Sub

Public Function rec_found(ByRef sRS As ADODB.Recordset, ByVal sField As String, ByVal sFindText As String, Optional ByVal dd As String) As Boolean
'-Move the recordset to the first record
sRS.Requery '-Use this instead of movefirst so that new record added can be used immediately
'Search the record
sRS.Find sField & " = '" & sFindText & "'"
'-Verify if the search string was found or not
If sRS.EOF Then
    rec_found = False
Else
    rec_found = True
End If
End Function

'function to centre form on the screen
Public Function CentreForm(frm As Form)
Dim x, y
        
x = (Screen.Width - frm.Width) / 2
y = (Screen.Height - frm.Height) / 2

frm.Move x, y

End Function

'function to determine whether the specified path is a file or a folder
Public Function IsFile(ByVal ObjName As String) As Boolean
If InStr(1, ObjName, ".") < 1 Then
IsFile = False
Else
IsFile = True
End If
End Function

'checks for the existence of a file, given a path
Public Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    ' If the string is quoted, remove the quotes.
    '
    strPathName = strUnQuoteString(strPathName)
    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = gstrSEP_DIR Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function

Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE And Right$(strQuotedString, 1) = gstrQUOTE Then
        '
        ' It's quoted.  Get rid of the quotes.
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    strUnQuoteString = strQuotedString
End Function

Public Function LockTask()
reg.SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 1

End Function


Public Function UnLockTask()
reg.SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0

End Function


