VERSION 5.00
Begin VB.Form Splash 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start!"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   3240
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      Picture         =   "splash.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   5250
      TabIndex        =   1
      Top             =   0
      Width           =   5250
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Skin"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Graphics Map"
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   16
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Feet:"
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Legs:"
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Torso:"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Hands:"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Belt:"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Arms:"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Head:"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Splash.Hide
GameForm.Show
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
Text1(0).Text = App.Path & "\graphics\head0.gif"
Text1(1).Text = App.Path & "\graphics\body-hands.gif"
Text1(2).Text = App.Path & "\graphics\body-arms.gif"
Text1(3).Text = App.Path & "\graphics\body-belt.gif"
Text1(4).Text = App.Path & "\graphics\body-torso.gif"
Text1(5).Text = App.Path & "\graphics\body-legs.gif"
Text1(6).Text = App.Path & "\graphics\body-feet.gif"
Text1(7).Text = App.Path & "\graphics\characterinfo.map"
Text1(8).Text = App.Path & "\graphics\gui.map"
End Sub
