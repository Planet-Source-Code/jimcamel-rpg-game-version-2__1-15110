VERSION 5.00
Begin VB.Form Editform 
   Caption         =   "Edit Form"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Change walkthroughable"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Add Tile"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Editform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = LoadPicture(App.Path & "/grass.bmp")
Picture1.AutoRedraw = False
Picture1.Refresh
If Button = 1 Then
startx = CInt(X / 16) * 16
starty = CInt(Y / 16) * 16
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Refresh
X1 = CInt(X / 16) * 16
Y1 = CInt(Y / 16) * 16
If Button = 1 Then
Picture1.Line (startx, starty)-(X1, Y1), , B
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tempx
Dim tempy
Picture1.Refresh
Picture1.AutoRedraw = True
If startx > X1 Then tempx = startx: startx = X1: X1 = tempx
If starty > Y1 Then tempy = starty: starty = Y1: Y1 = tempy
Picture1.Line (startx, starty)-(X1, Y1), , B
End Sub
