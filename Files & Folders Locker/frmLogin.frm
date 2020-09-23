VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Wildfox Files & Folders Locker ver 1.0.0"
   ClientHeight    =   3405
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2011.787
   ScaleMode       =   0  'User
   ScaleWidth      =   4337.929
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2640
      Top             =   2880
   End
   Begin VB.CommandButton testme 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   2445
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2445
   End
   Begin VB.Label lblattemptcout 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblattemptcount 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
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
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1305
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Dim C_Attempt As String
Dim logtime As Date
Dim strtime As Date
Dim medd As String
Public logintime As Date


Private Sub Command1_Click()
Me.Hide
frmmain.mnutray_Click

End Sub

Private Sub Form_Initialize()
InitCommonControls
Load frmmain

LockTask
End Sub

Private Sub Form_Load()
App.Title = ""
If App.PrevInstance = True Then
MsgBox "A previous instance of this application is already running", vbCritical, strAppTitle
UnloadAllForms
End If

If FileExists(App.Path & "\" & "loadhigh.exe") = False Then
MsgBox "This program has detected that critical program files have been deleted or modified by an unknown agent!" & vbCrLf & "Please reinstall the application immediately. Remember your files and folders are no longer protected!" & vbCrLf & "Protection will resume after reinstallation. The application will now terminate.", vbCritical, "System Integrity Compromised"
UnloadAllForms
Exit Sub
End If

CentreForm Me

Dim KeyAscii As KeyCodeConstants
Set cn = New ADODB.Connection
cn.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\LockInfo.mdb;Persist Security Info=False;Jet OLEDB:Database Password = wildfox")
If rs_user.State = adStateOpen Then
rs_user.Close
End If
rs_user.Open "select * from tbl_securityconfig", cn, adOpenDynamic, adLockOptimistic

If rs_secuser.State = adStateOpen Then
rs_secuser.Close
End If

C_Attempt = 3
lblattemptcount.Caption = "Attempts Remaining: " & C_Attempt



If KeyAscii = 13 Then
testme_click
End If
medd = App.Path & "\LockInfo.mdb"

End Sub


Private Sub testme_click()
'Verify the fields if empty
'testme_click
If txtusername.Text = "" Then txtusername.SetFocus: Exit Sub
If txtpassword.Text = "" Then txtpassword.SetFocus: Exit Sub
'Check if the User Name is valid
If rec_found(rs_user, "UserName", txtusername.Text) = False Then
    C_Attempt = C_Attempt - 1
    lblattemptcount.Visible = True
    lblattemptcount.Caption = "Attempts Remaining " & C_Attempt
    
    If C_Attempt < 1 Then
        'MsgBox "Maximum login attempts has been reached." & vbCrLf & "Possible attempted Security breech detected " & vbCrLf & "This application will now Terminate.", vbCritical, strAppTitle
         
         C_Attempt = 3
         txtusername.Text = ""
         txtpassword.Text = ""
         lblattemptcount.Caption = ""
         frmmain.mnutray_Click
         Command1_Click
  Else:
    lblattemptcount.Visible = True
    MsgBox "The User Name you entered is not valid." & vbCrLf & "Please try again." & vbCrLf & vbCrLf & "Warning: You only have " & C_Attempt & " attempts.", vbCritical, strAppTitle
    lblattemptcount.Caption = "Attempts Remaining " & C_Attempt
    SendKeys "{home}"
    txtusername.SetFocus
  
    End If
    
    Exit Sub
End If
'Check if the Password is valid
If txtpassword.Text <> rs_user.Fields("Password") Then
    C_Attempt = C_Attempt - 1
    lblattemptcount.Visible = True
    lblattemptcount.Caption = "Attempts Remaining: " & C_Attempt

    If C_Attempt < 1 Then
        'MsgBox "Maximum login attempts has been reached." & vbCrLf & "Possible attempted Security breech detected" & vbCrLf & "This application will now Terminate.", vbCritical, strAppTitle
        
        C_Attempt = 3
        txtusername.Text = ""
        txtpassword.Text = ""
        lblattemptcount.Caption = ""
        frmmain.mnutray_Click
        Command1_Click
    Else
        lblattemptcount.Visible = True
        MsgBox "You did NOT enter the Correct Password." & vbCrLf & "Please try again." & vbCrLf & vbCrLf & "Warning: You only have " & C_Attempt & " attempt.", vbCritical, strAppTitle
        lblattemptcount.Caption = "Attempts Remaining: " & C_Attempt
        txtpassword.SetFocus
        SendKeys "{home}"
        
    
    End If
    
    
    Exit Sub
End If


'Copy the Username and log-time to variable
With CurrentUser
     .username = txtusername.Text
     .password = txtpassword.Text
     .user_timelogin = Time
     .usergroup = rs_user.Fields("usergroup")
End With
 reg.SaveString HKEY_LOCAL_MACHINE, "Software\Wildfox Multimedia\Wildfox Locker\General", "Currrent User", CurrentUser.username
 reg.SaveString HKEY_LOCAL_MACHINE, "Software\Wildfox Multmedia\Wildfox Locker\General", "timelogin", Time

logintime = Time
Unload frmLogin
frmmain.Show

If Not frmmain.sbFox.Panels.Item(2).Text = CurrentUser.username Then
frmmain.sbFox.Panels.Item(2).Text = CurrentUser.username
frmmain.sbFox.Panels.Item(6).Text = Format(Date, "dd/mm/yyyy")
frmmain.sbFox.Panels.Item(4).Text = CurrentUser.user_timelogin
frmmain.MSHFlexGrid1.Visible = False
End If
frmmain.List1.Visible = False

End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
testme_click
End If

End Sub
Public Function rec_found(ByRef sRS As ADODB.Recordset, ByVal sField As String, ByVal sFindText As String) As Boolean
'Move the recordset to the first record
sRS.Requery '-Use this instead of movefirst so that new record added can be used immediately
'Search the record
sRS.Find sField & " = '" & sFindText & "'"
'Verify if the search string was found or not
If sRS.EOF Then
    rec_found = False
Else
    rec_found = True
End If
End Function

Private Sub txtUserName_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
testme_click
End If

End Sub

