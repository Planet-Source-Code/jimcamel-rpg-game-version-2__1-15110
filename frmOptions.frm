VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   2550
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sound Effects"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Music"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then Work.StopMIDI
If Check1.Value = 1 Then
Work.PlayMIDI
If musicinit = False Then Work.iniMIDI
Work.LoadMIDI App.Path & "\music.mid"
Work.PlayMIDI
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then Work.Stopwave
If Check2.Value = 1 And soundinit = False Then Work.initsound: Work.LoadWave 1, App.Path & "\tsteps.wav"
End Sub

Private Sub Command1_Click()
SaveSetting "JimCamel", "RPGame", "Sound", Check2.Value
SaveSetting "JimCamel", "RPGame", "Music", Check1.Value
frmOptions.Hide
End Sub

Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
End Sub
