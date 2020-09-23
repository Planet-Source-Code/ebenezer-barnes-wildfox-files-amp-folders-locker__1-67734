VERSION 5.00
Begin VB.Form frmusers 
   Caption         =   "Users"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmusers.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox frainfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   50
      Picture         =   "frmusers.frx":19D8
      ScaleHeight     =   3225
      ScaleWidth      =   5625
      TabIndex        =   13
      Top             =   50
      Width           =   5655
      Begin VB.TextBox txtcomments 
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtfirstname 
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   1290
         Width           =   2415
      End
      Begin VB.TextBox txtlastname 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   1695
         Width           =   2415
      End
      Begin VB.TextBox txtusername 
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cmbusergroup 
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Text            =   "Users"
         Top             =   2115
         Width           =   2415
      End
      Begin VB.TextBox txtpassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtverifypassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblverifypassword 
         BackColor       =   &H000000A8&
         Caption         =   "Verify Password:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblpassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblcomments 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
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
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lbllname 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblfname 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblusername 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblusergroup 
         BackStyle       =   0  'Transparent
         Caption         =   "User Group"
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.PictureBox picnav 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   50
      Picture         =   "frmusers.frx":33B0
      ScaleHeight     =   705
      ScaleWidth      =   5625
      TabIndex        =   10
      Top             =   3280
      Width           =   5660
      Begin VB.CommandButton Command5 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   100
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Prev. Record"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   100
         Width           =   1080
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Next Record"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   100
         Width           =   1080
      End
      Begin VB.CommandButton Command1 
         Caption         =   "First Record"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   100
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Last Record"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   4
         Top             =   100
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   4380
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox picedit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   5640
      Picture         =   "frmusers.frx":4D88
      ScaleHeight     =   3225
      ScaleWidth      =   1305
      TabIndex        =   0
      Top             =   50
      Width           =   1335
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdnew 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()

LockFrames
rs_user.CancelUpdate
cmdrefresh_Click
UnLockControl
UnSetConfirm

End Sub

Private Sub cmddelete_Click()
If CurrentUser.usergroup = "Administrators" Then

 If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo + vbDefaultButton2, "Delete Record") <> vbYes Then
     MsgBox "Delete operation has been canceled by user", vbInformation, strAppTitle
     Exit Sub
  End If
  With rs_user
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
Else
    MsgBox "Insufficient security previlages for the current user!" & vbCrLf & "Username: " & CurrentUser.username & vbCrLf & "Security Group Membership: " & CurrentUser.usergroup, vbExclamation, strAppTitle
End If


End Sub

Private Sub cmdedit_Click()

UnLockFrames
EditFlag = True
picnav.Enabled = False
LockControl

End Sub

Private Sub cmdnew_Click()
On Error GoTo AddErr

If CurrentUser.usergroup = "Administrators" Then
    rs_user.AddNew
    
    UnLockFrames
    LockControl
    FromAddnew = True
    AddnewFlag = True
Else
    MsgBox "Insufficient security previlages for the current user!" & vbCrLf & "Username: " & CurrentUser.username & vbCrLf & "Security Group Membership: " & CurrentUser.usergroup, vbExclamation, strAppTitle
End If

AddErr:
If Err.Number = -2147217887 Then
MsgBox "Account already exits in the database, Please try again.", vbInformation, "Wildfox Cafepro Server"
Exit Sub
End If
End Sub

Private Sub cmdrefresh_Click()
LoadRecords
LockFrames
End Sub

Private Sub Command1_Click()
rs_user.MoveFirst
End Sub

Private Sub Command2_Click()
MovetoPrevious
End Sub

Private Sub Command3_Click()
MoveRoutine
End Sub

Private Sub Command4_Click()
MovetoNext
End Sub

Private Sub Form_Initialize()
With cmbusergroup
.AddItem "Administrators"
.AddItem "Users"
End With
End Sub

Private Sub Form_Load()
ConnectToDB
CentreForm Me
LoadRecords
LockFrames

FromAddnew = False
EditFlag = False
AddnewFlag = False
frainfo.Picture = Me.Picture
picnav.Picture = Me.Picture
picedit.Picture = Me.Picture
End Sub





Public Sub LoadRecords()
Set rs_user = Nothing

ConnectToDB
If CurrentUser.usergroup = "Administrators" Then
rs_user.Open "SELECT * FROM tbl_securityconfig", cn, adOpenDynamic, adLockOptimistic
Else
rs_user.Open "SELECT * FROM tbl_securityconfig where username=" & "'" & CurrentUser.username & "'", cn, adOpenDynamic, adLockOptimistic
End If

With txtusername
Set .DataSource = rs_user
.DataField = "username"
End With

With txtpassword
Set .DataSource = rs_user
.DataField = "password"
End With

With txtfirstname
Set .DataSource = rs_user
.DataField = "fname"
End With

With txtlastname
Set .DataSource = rs_user
.DataField = "lname"
End With

With txtcomments
Set .DataSource = rs_user
.DataField = "comments"
End With

With cmbusergroup
Set .DataSource = rs_user
.DataField = "usergroup"
End With
End Sub


Public Sub MoveRoutine()
On Error GoTo TrapMe
rs_user.MoveLast

TrapMe:
If Err.Number = -2147217842 Then
'do othing
End If
End Sub

Public Function MovetoPrevious()
On Error GoTo moveerr

If Not rs_user.BOF Then
rs_user.MovePrevious
If rs_user.BOF Then
rs_user.MoveFirst
End If
End If

moveerr:
If Err.Number = -2147467259 Then
Exit Function
End If

End Function
Public Function MovetoNext()

On Error GoTo moveerr

If Not rs_user.EOF Then
rs_user.MoveNext
If rs_user.EOF Then
rs_user.MoveLast
End If
End If

moveerr:
If Err.Number = -2147467259 Then
Exit Function
End If

End Function


Private Sub cmdupdate_Click()
Dim rr As New ADODB.Recordset
Dim sc As New ADODB.Recordset
 rr.Open "Select * from tbl_securityconfig", cn, 2, 3
If Not EditFlag = True Then
       
    If rec_found(rr, "username", txtusername.Text) = True Then
    MsgBox "Username already exists in Database. Please check and try again!", vbInformation, "Duplicate Username"
    txtusername.SetFocus: SendKeys "{Home}+{End}"
    Set rr = Nothing
    Exit Sub
    End If
    
    If txtpassword.Text <> txtverifypassword.Text Then
    MsgBox "The entered passwords do not match", vbExclamation, strAppTitle
    Exit Sub
    End If
End If

If Trim(txtusername.Text) = vbNullString Then
    MsgBox "Invalid Username. Please check and try again!", vbInformation, "Invalid Username"
    txtusername.SetFocus: SendKeys "{Home}+{End}"
    Exit Sub
End If

If Trim(txtpassword.Text) = vbNullString Then
    MsgBox "Invalid password. Please check and try again!", vbInformation, "Invalid Password"
    txtpassword.SetFocus: SendKeys "{Home}+{End}"
    Exit Sub
End If

If Not cmbusergroup.Text = vbNullString Then
    Select Case cmbusergroup.Text
        Case "Administrators"
        'User belongs to the administrators user group
        
        Case "Users"
        'User belongs to the ordinary users group
        
        Case Else
        MsgBox "Invalid users group. Please check and try again!", vbInformation, "Invalid User Group"
        cmbusergroup.SetFocus: SendKeys "{Home}+{End}"
    Exit Sub
End Select
End If


rs_user.UpdateBatch adAffectAll


If EditFlag = True Then
    MovetoNext
    MovetoPrevious
End If


FromAddnew = False
EditFlag = False
AddnewFlag = False

LockFrames
UnLockControl
UnSetConfirm
End Sub


Public Sub LockFrames()
frainfo.Enabled = False
cmdcancel.Enabled = False
cmdupdate.Enabled = False
End Sub

Public Sub UnLockFrames()
frainfo.Enabled = True
picedit.Enabled = True
cmdcancel.Enabled = True
cmdupdate.Enabled = True
End Sub


Public Sub UnLockControl()
cmdnew.Enabled = True
cmdedit.Enabled = True
cmddelete.Enabled = True
cmdrefresh.Enabled = True
picnav.Enabled = True
End Sub

Public Sub LockControl()
cmdnew.Enabled = False
cmdedit.Enabled = False
cmddelete.Enabled = False
cmdrefresh.Enabled = False
picnav.Enabled = False
End Sub


Private Sub txtpassword_LostFocus()
If FromAddnew = True Then
    SetConfirm
    txtverifypassword.SetFocus
End If
End Sub


Private Sub SetConfirm()
lblverifypassword.Visible = True
txtverifypassword.Visible = True
lblpassword.Visible = False
txtpassword.Visible = False
End Sub

Private Sub UnSetConfirm()
lblverifypassword.Visible = False
txtverifypassword.Visible = False
lblpassword.Visible = True
txtpassword.Visible = True
End Sub

