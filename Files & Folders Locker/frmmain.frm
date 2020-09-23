VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wildfox Files &  Folders Locker"
   ClientHeight    =   6645
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9240
   ControlBox      =   0   'False
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmain.frx":0000
   ScaleHeight     =   6645
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4440
      Left            =   20
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   7832
      _Version        =   393216
      BackColor       =   168
      ForeColor       =   16777215
      Rows            =   20
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorBkg    =   168
      GridColor       =   12632256
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "Object Name|Owner|Protection Date|Last Unlock Date"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   0
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H000000A8&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   8775
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H000000A8&
      ForeColor       =   &H80000005&
      Height          =   3210
      Left            =   4440
      TabIndex        =   9
      Top             =   1320
      Width           =   4575
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H000000A8&
      ForeColor       =   &H80000005&
      Height          =   3240
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   4095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H000000A8&
      Height          =   3765
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.CommandButton cmdLockProtected 
      Caption         =   "Lock Protected Objects"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000A8&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   4890
      Width           =   8775
   End
   Begin MSComctlLib.StatusBar sbFox 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   6150
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Text            =   "Username:"
            TextSave        =   "Username:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Text            =   "Timelogin:"
            TextSave        =   "Timelogin:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdProtectedObjects 
      Caption         =   "Show Protected Objects"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdlock 
      Caption         =   "Lock"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnutray 
         Caption         =   "&Minimize to Tray"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuusers 
         Caption         =   "&Users"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "&Log Off"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnutaskmgr 
         Caption         =   "&Enable Task Manager"
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   +^{F4}
      End
   End
   Begin VB.Menu mnupop 
      Caption         =   "Mnupop"
      Begin VB.Menu mnupshow 
         Caption         =   "&Show"
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu mnupabout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuPFF 
      Caption         =   "&PFF"
      Begin VB.Menu mnuunlock 
         Caption         =   "Unlock"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuunprotect 
         Caption         =   "Un&protect"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpl 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnucontact 
         Caption         =   "Contact Us"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:    Ebenezer Essel Barnes
'Address:   P.O. Box MD 53, Madina-Accra, Ghana
'email:     docebarnes@yahoo.com
'Cell Phone: +233243947960

'Based on the file lock program submitted by Pierre Aoun on PSC
'Added a lot of modifications to make it more useful
'This is a program I intend to release commercially in the near future but have submitted it here for peer review
'It is still under development and therefore there may be objects included that are not being used at present.


'Special thanks to Pierre Aoun for the very useful insight on the
'file or folder handle locking.

'To those coders who always selflessly share code on PSC, this is my
'way of saying thanx as I have learned a lot from you guys.

'Please to all coders and folks on PSC: your suggestions, comments and votes
'are badly needed so as to help me improve on the program and make it more
'secure. You can write to me or drop a mail if you think of any feature that needs
'to be implemented or you find sections of code that can be made more useful. Thank you.


'Features include:
'1. stealth startup? (curtesy Dr. Y,  mosibatzadeh@yahoo.com)
'2. Multi-user environment
'3. view protected objects per user basis
'4. Tempering proof? (program shuts down automatically on files modification or deletion).


'Tested on win2000, winXP, am sure will work under win98 with little or no troubles
'well you can try that since I don't have a win98 machine for testing.

'One nagging problem: I tried to use the hide process procedures as outlined by Islam Adel (Breakthrough !! UPDATED - Completely hiding a process from the task manager in 9x and NT!) but I always get the class not initialised error.
'I wish some one can test this program against mine and make it work, in that case the app will be truely invincible.
'If you are able to implement that, I wish you send the modified code by mail to me as I intend to develop similar apps on this line. Thank you.

Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal PassZero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal PassZero As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()



Private datPrimaryRS As ADODB.Recordset

Dim hDir As Long
Dim i As Integer
Dim rs_jj As New ADODB.Recordset
Public DirPath, FilePath As String
Public Mscript As New Scripting.FileSystemObject
Public PathDir As String

'sub to lock the file or folder
Private Sub cmdlock_Click()
On Error GoTo ProErr
Dim a, b, c, d, ObjType As String
Dim ans, e As Integer


If Text1.Text = "" Then MsgBox "No object selected for the current operation", vbCritical, strAppTitle: Exit Sub

If IsFile(Text1.Text) = True Then
ObjType = "file"
Else
ObjType = "folder"
End If


 'Check to see if the user is trying to protect a root folder (eg. c:\)
 
    a = Mscript.GetParentFolderName(Text1.Text)
    b = Mscript.GetSpecialFolder(SystemFolder)
    c = Mscript.GetSpecialFolder(WindowsFolder)
    d = Mscript.GetSpecialFolder(TemporaryFolder)
    'e = InStr(1, Text1.Text, "system32")
    
     If a = "" Then
     MsgBox "Invalid operation requested! Program does not Lock Root Folders", vbCritical, "Illegal Request"
     Exit Sub
     End If
     
 'Check whether the user is trying to protect a key system folder (eg. c:\windows)
     
     Select Case Text1.Text
     Case b
          MsgBox "Invalid operation requested! Program does not Lock Key System Folders" & vbCrLf & "Object Requested: " & b, vbCritical, "Illegal Request"
          Exit Sub
    Case c
          MsgBox "Invalid operation requested! Program does not Lock Key System Folders" & vbCrLf & "Object Requested: " & c, vbCritical, "Illegal Request"
          Exit Sub
    Case d
          MsgBox "Invalid operation requested! Program does not Lock Key System Folders" & vbCrLf & "Object Requested: " & d, vbCritical, "Illegal Request"
          Exit Sub
    End Select

  
    
ans = MsgBox("Do you want to protect the current " & ObjType & ".", vbInformation + vbYesNo + vbDefaultButton2, strAppTitle)

If ans = vbYes Then

     'hDir would hold the handle to the file or folder
        PathDir = Text1.Text
        hDir = CreateFile(PathDir, FILE_LIST_DIRECTORY, File_Share_Flag, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, ByVal 0&)
       
     'This piece code was put in place to check if we can read the file if not then its been locked
     'I think there might be a better way of doing this but I will maintain this for now.
        If InStr(1, PathDir, ".") < 1 Then
        Mscript.CopyFolder PathDir, "c:\locktemp"
        Else
        Mscript.CopyFile PathDir, "c:\LockTemp"
        End If
 Else
 MsgBox "Protect operation was canceled by user", vbInformation, strAppTitle
 End If
 
 
ProErr:
'If we get an error telling us we were unable to read the file then the file or folder has been locked successfully
'We then write the details to the database
    If Err.Number = 70 Then
    MsgBox "Protect Operation Completed Successfully", vbInformation, "Lock Successfull"
    WriteToDatabase
    End If
    
    
End Sub


Public Sub cmdLockProtected_Click()

'clear the listbox
List1.Clear

'close the recordset if its open to prevent errors
If rs_secuser.State = adStateOpen Then rs_secuser.Close
If rs_lockinfo.State = adStateOpen Then rs_lockinfo.Close

'open the recordset and add the records to the listbox

rs_lockinfo.Open "SELECT * FROM lockinginfo", cn, 3, 3

If Not rs_lockinfo.EOF Then
rs_lockinfo.MoveFirst
End If
Do While Not rs_lockinfo.EOF

List1.AddItem rs_lockinfo.Fields("objname")
rs_lockinfo.MoveNext
Loop
rs_lockinfo.Close


'lock the protected files read from the database

i = i + 1
For i = 0 To List1.ListCount
Label1.Caption = List1.List(i)


PathDir = Label1.Caption
hDir = CreateFile(PathDir, FILE_LIST_DIRECTORY, File_Share_Flag, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, ByVal 0&)

If rs_jj.State = adStateOpen Then rs_jj.Close
rs_jj.Open "SELECT * FROM lockinginfo where objname =" & "'" & PathDir & "'", cn, 3, 3

  If rec_found(rs_jj, "objname", PathDir) = True Then
  With rs_jj
    .Fields("fhandle") = hDir
    .Update
    End With
    End If
    

Next i


End Sub



Private Sub cmdProtectedObjects_Click()


List1.Clear
If cmdProtectedObjects.Caption = "Show Protected Objects" Then
    List1.Visible = True
    cmdProtectedObjects.Caption = "Hide Protected Objects"
    cmdlock.Enabled = False
    'Frame1.Caption = "Protected Files and Folders"
    
    If rs_lockinfo.State = adStateOpen Then rs_lockinfo.Close
    
    'open the recordset and add the records to the listbox
        If Not CurrentUser.usergroup = "Administrators" Then
                rs_lockinfo.Open "SELECT * FROM lockinginfo where username=" & "'" & CurrentUser.username & "'", cn, 3, 3
           Else
                rs_lockinfo.Open "SELECT * FROM lockinginfo", cn, 3, 3
                End If
    
    
        If Not rs_lockinfo.EOF Then
            rs_lockinfo.MoveFirst
        End If
        
        Do While Not rs_lockinfo.EOF
    
    List1.AddItem rs_lockinfo.Fields("objname")
    rs_lockinfo.MoveNext
    Loop
    MSHFlexGrid1.Visible = True
    LoadDetails


Else
    List1.Visible = False
    MSHFlexGrid1.Visible = False
    cmdProtectedObjects.Caption = "Show Protected Objects"
    cmdlock.Enabled = True
    'Frame1.Caption = "File or Folder Selection"
End If


End Sub



Private Sub Dir1_Change()
On Error GoTo A1:
    DirPath = Dir1.Path
    Text1.Text = DirPath

File1.Path = Dir1.Path
    Exit Sub
A1:
    MsgBox "Folder Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Drive1_Change()
Text1.Text = ""
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub File1_Click()
Dim a, b As String

FilePath = File1.FileName
a = Dir1.Path & "\" & FilePath
b = Mscript.GetParentFolderName(a)

If Len(b) = 3 Then
    Text1.Text = Dir1.Path & FilePath
Else
    Text1.Text = a
End If

End Sub

Private Sub Form_Initialize()
InitCommonControls
Hasrun = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyT Then
UnLockTask
End If

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
UnloadAllForms
End If

Hasrun = False

C_Attempt = 3
ConnectToDB
CentreForm Me
File_Share_Flag = 0 'if =FILE_SHARE_READ then read only (for example)
    
CreateFolder
cmdLockProtected_Click
i = -1
mnuPFF.Visible = False
mnupop.Visible = False
List1.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' Unloadmode codes
'     0 - Close from the control-menu box or Upper right "X"
'     1 - Unload method from code elsewhere in the application
'     2 - Windows Session is ending
'     3 - Task Manager is closing the application
'     4 - MDI Parent is closing
' ---------------------------------------------------------------------------
  Select Case UnloadMode
  
         Case 1: UnloadAllForms
         Case 2: UnloadAllForms
         Case 3: mnutray_Click
          
 End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim i As Integer
    Dim fn As File
    Dim s As Folder
    
    Set s = Mscript.GetFolder("c:\LockTemp")
    
    For Each fn In s.Files
    fn.Delete
    Next
      
End Sub


Sub CreateFolder()
    If Mscript.FolderExists("c:\LockTemp") = False Then
       Mscript.CreateFolder ("C:\LockTemp")
    Else
       Exit Sub
    End If
End Sub

Public Sub WriteToDatabase()
Dim s As String
'sub to write the file or folder details to the database

'connect tot the database
ConnectToDB
If rs_secuser.State = adStateOpen Then rs_secuser.Close
rs_secuser.Open "Select * from lockinginfo", cn, adOpenDynamic, adLockOptimistic

'add the protected file or folder to the database

s = Format(Date, "dd/mm/yyyy ") & Time
With rs_secuser
.AddNew
.Fields("objname") = PathDir
.Fields("isfile") = IsFile(PathDir)
.Fields("Fhandle") = hDir
.Fields("username") = CurrentUser.username
.Fields("protectiondate") = s
.Fields("lastunlockdate") = s
.Update
End With
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
PopupMenu mnuPFF, vbPopupMenuRightButton
End If

End Sub

Private Sub mnuexit_Click()
If CurrentUser.usergroup = "Administrators" Then
    Dim ans As Integer
    ans = MsgBox("Are you sure you want to exit the program?" & vbCrLf & "This will end protection of all Files and Folders currently protected." & vbCrLf & "Protection of all Files and Folders will resume once the program has been restarted.", vbCritical + vbYesNo, strAppTitle)
    
    If ans = vbYes Then
        UnLockTask
        UnloadAllForms
    Else
        Exit Sub
    End If
Else
    MsgBox "Insufficient security previlages for the current user!" & vbCrLf & "Username: " & CurrentUser.username & vbCrLf & "Security Group Membership: " & CurrentUser.usergroup, vbExclamation, strAppTitle
End If

End Sub



Public Sub SysTrayMouseEventHandler()
SetForegroundWindow Me.hwnd
PopupMenu mnupop, vbPopupMenuRightButton
End Sub


Private Sub mnulogoff_Click()
If Not MsgBox("Are you sure you want to logoff", vbInformation + vbYesNo, strAppTitle) <> vbYes Then
    frmmain.Hide
    frmLogin.Show vbModal
Else
    MsgBox "Logoff operation aborted by user", vbInformation, strAppTitle
End If

End Sub

Private Sub mnupshow_Click()
UnhookForm
frmmain.Visible = False
frmLogin.Show vbModal
End Sub

Private Sub mnutaskmgr_Click()
UnLockTask
End Sub

Public Sub mnutray_Click()
On Error GoTo TrayErr
Hooks Me.hwnd   ' Set up our handler
AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, "Wildfox Files & Folders Locker"
Me.Hide

TrayErr:
If Err.Number = 402 Then
'ignore error
End If


End Sub

Private Sub mnuunlock_Click()
'sub to unlock a locked file
UnlockUsingGrid
End Sub

Private Sub mnuunprotect_Click()
'sub to unprotect the file or folder
UnprotectUsingGrid

End Sub

Public Function UnhookForm()
'function to remove the form's icon from the system tray
Unhook    ' Return event control to windows
Me.Show
RemoveIconFromTray
End Function

Private Sub mnuusers_Click()
frmusers.Show vbModal
End Sub


Private Sub LoadDetails()
ConnectToDB
    Dim sConnect As String
    Dim sSQL As String
    Dim dfwConn As ADODB.Connection
    Dim i As Integer
    Dim j As Integer
    Dim m_iMaxCol As Integer

    ' set strings
    If Not CurrentUser.usergroup = "Administrators" Then
    sSQL = "SELECT lastunlockdate,objname,protectiondate,username,isfile FROM LockingInfo where username=" & "'" & CurrentUser.username & "'"
    
    Else
    sSQL = "SELECT lastunlockdate,objname,protectiondate,username, isfile FROM LockingInfo"

    End If


    ' open connection
    Set dfwConn = New Connection

    ' create a recordset using the provided collection
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly


    Set MSHFlexGrid1.DataSource = datPrimaryRS

'this piece of code is to check the value of the field "isfile"
'and represent it in the grid as a file or folder
'nicer than -1 or 0, don't you think?
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim ki As Integer
Dim kkk As String

For ki = MSHFlexGrid1.FixedRows To MSHFlexGrid1.Rows - 1
MSHFlexGrid1.Row = ki
kkk = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
If kkk = "True" Then
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = "File"
Else
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4) = "Folder"
End If
Next ki
'end of code
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    With MSHFlexGrid1

        .Redraw = False
        ' place the columns in the right order
        .ColData(1) = 0
        .ColData(3) = 1
        .ColData(2) = 2
        .ColData(0) = 3
        .ColData(4) = 4
        
       

        ' loop to re-order the columns
        For i = 0 To .Cols - 1
            m_iMaxCol = i                   ' find the highest value starting from this column
            For j = i To .Cols - 1
                If .ColData(j) > .ColData(m_iMaxCol) Then m_iMaxCol = j
            Next j
            .ColPosition(m_iMaxCol) = 0     ' move the column with the max value to the left
        Next i

        ' modify column's headers
        .TextMatrix(0, 0) = "Object Name"
        .TextMatrix(0, 4) = "Type"
        .TextMatrix(0, 1) = "Owner"
        .TextMatrix(0, 2) = "Protection Date"
        .TextMatrix(0, 3) = "Last Unlock Date"
        ' set grid's column widths
        .ColWidth(0) = 4000
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 900

        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' make header bold
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True
        .CellFontSize = 10

        ' grey every other row
        For i = .FixedRows + 1 To .Rows - 1 Step 2
            .Row = i
            .Col = .FixedCols
            .ColSel = .Cols() - .FixedCols - 1
            .CellBackColor = &HC0C0C0   ' light grey
            
        Next i

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyT Then
UnLockTask
End If

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
PopupMenu mnuPFF, vbPopupMenuRightButton
End If
End Sub

Private Sub UnlockUsingGrid()
'sub to unlock a locked file
Dim strObjname As String
strObjname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)

If strObjname = "" Then MsgBox "No object selected for the current operation", vbCritical, strAppTitle: Exit Sub
If strObjname = "Object Name" Then MsgBox "No object selected for the current operation", vbCritical, strAppTitle: Exit Sub

Dim ss As New ADODB.Recordset
Dim rr As New ADODB.Recordset
Dim FileHandle As Integer
Dim ans As Integer
Dim ObjType As String
Dim t As String

If IsFile(strObjname) = True Then
ObjType = "file"
Else
ObjType = "folder"
End If

If Not strObjname = vbNullString Then
ans = MsgBox("Are you sure you want to unlock this " & ObjType & "?" & vbCrLf & "Object Name: " & strObjname, vbInformation + vbYesNo + vbDefaultButton2, strAppTitle)

If ans = vbYes Then

'connect to the database
ConnectToDB
ss.Open "SELECT * FROM lockinginfo", cn, 3, 3

'check if the specified file or folder exits
'if true close the handle to the file to unlock it.
If rec_found(ss, "objname", strObjname) = True Then
rr.Open "SELECT * FROM lockinginfo where objname=" & "'" & strObjname & "'", cn, 3, 3
FileHandle = rr.Fields("fhandle")
CloseHandle FileHandle

t = Format(Date, "dd/mm/yyyy ") & Time

With rr
.Fields("lastunlockdate") = t
.Update
End With
End If

Else
MsgBox "Unlock operation was canceled by user", vbInformation, strAppTitle

End If
Else
MsgBox "No object selected for the unlock operation.", vbCritical, strAppTitle
Exit Sub
End If
End Sub


Private Sub UnprotectUsingGrid()
Dim strObjname As String
strObjname = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)

If strObjname = "" Then MsgBox "No object selected for the current operation", vbCritical, strAppTitle: Exit Sub

Dim ans As Integer
Dim sr As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim ObjType As String

If IsFile(strObjname) = True Then
ObjType = "file"
Else
ObjType = "folder"
End If

'ask for action confirmation
ans = MsgBox("Are you sure you want to PERMANENTLY stop protecting this " & ObjType & "?" & vbCrLf & "Object Name: " & strObjname & vbCrLf & "Remember to Unlock the " & ObjType & " first before unprotecting it." & vbCrLf & "Do you still want to carry out the current operation?", vbCritical + vbYesNo + vbDefaultButton2, strAppTitle)

'if the response is yes then delete the record from the database
If ans = vbYes Then
    rs.Open "SELECT * FROM lockinginfo", cn, 3, 3

    If rec_found(rs, "objname", strObjname) = True Then
        sr.Open "SELECT * FROM lockinginfo where objname=" & "'" & strObjname & "'", cn, 3, 3
        
        sr.Delete adAffectCurrent
    End If
    Else
    MsgBox "Unprotect operation was canceled by user", vbInformation, strAppTitle
    Exit Sub
End If
rs.Close
sr.Close
End Sub

