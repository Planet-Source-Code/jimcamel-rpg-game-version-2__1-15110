VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2460
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1450.633
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AeonNexus.GetServer GetServer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmLogin.frx":0000
      Left            =   1290
      List            =   "frmLogin.frx":0002
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   915
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Save Details"
      Height          =   252
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   1290
      TabIndex        =   2
      Text            =   "5558"
      Top             =   1305
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   2040
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label2 
      Caption         =   "&Port:"
      Height          =   252
      Left            =   105
      TabIndex        =   9
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "&Server:"
      Height          =   252
      Left            =   105
      TabIndex        =   8
      Top             =   930
      Width           =   1092
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim serverip As String
If InStr(1, LCase(Combo1.Text), "main") > 0 And InStr(1, LCase(Combo1.Text), "server") > 0 Then
serverip = GetServer1.DownloadText("http://an.egn3d.com/serverip.txt")
Do: DoEvents: Loop While serverip = ""
Else
serverip = Combo1.Text
End If
MsgBox serverip
On Error Resume Next
If Check1.Value = 1 Then
    SaveSetting "JimCamel", "RPGame", "Username", txtUserName.Text
    SaveSetting "JimCamel", "RPGame", "Password", txtPassword.Text
    SaveSetting "JimCamel", "RPGame", "Hostname", Combo1.Text
    SaveSetting "JimCamel", "RPGame", "IPAddress", Text2.Text
End If
GameForm.Winsock1.RemotePort = frmLogin.Text2.Text
GameForm.Winsock1.RemoteHost = serverip
GameForm.Winsock1.Connect
End Sub

Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
On Error GoTo 1
Dim User As String
Dim Pass As String
Dim Host As String
Dim IP As String
User = GetSetting("JimCamel", "RPGame", "Username")
Pass = GetSetting("JimCamel", "RPGame", "Password")
Host = GetSetting("JimCamel", "RPGame", "Hostname")
IP = GetSetting("JimCamel", "RPGame", "IPAddress")
If User <> "" Then
Let txtUserName.Text = User
Else: txtUserName.Text = ""
End If
If Pass <> "" Then
Let txtPassword.Text = Pass
Else: txtPassword.Text = ""
End If
If Host <> "" Then
Combo1.AddItem ("Main Server")
Combo1.AddItem (Host)
Else:
Combo1.AddItem ("Main Server")
End If
Combo1.ListIndex = 0
If IP <> "" Then
Let Text2.Text = IP
Else: Text2.Text = "6005"
End If
Exit Sub
1
End Sub
