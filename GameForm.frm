VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form GameForm 
   BackColor       =   &H80000004&
   Caption         =   "Aeon Nexus"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   DrawStyle       =   6  'Inside Solid
   FillStyle       =   7  'Diagonal Cross
   BeginProperty Font 
      Name            =   "Tempus Sans ITC"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GameForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   7200
      Width           =   9615
   End
   Begin VB.PictureBox OpenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox DXPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillStyle       =   7  'Diagonal Cross
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   1800
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   840
         Top             =   120
      End
      Begin VB.Timer Movetimer 
         Interval        =   10
         Left            =   1320
         Top             =   120
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1200
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   7560
      Visible         =   0   'False
      Width           =   9615
   End
End
Attribute VB_Name = "GameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DXPic_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then quitmenu = Not quitmenu: Bltbas.Blt
If KeyCode = vbKeyL Then
Load frmLogin
frmLogin.Show
frmLogin.SetFocus
End If
If KeyCode = vbKeyF7 Then frmplayerlist.Show
If KeyCode = vbKeyO Then frmOptions.Show: frmOptions.SetFocus
If KeyCode = vbKeyN And edit = False Then Playerdetails(0) = InputBox("Please enter your name", "Enter Name", "JimCamel"): SaveSetting "JimCamel", "RPGame", "Playername", Playerdetails(0): Bltbas.Blt: If connected = True Then Winsock1.SendData "chnm:" & Playerdetails(0) & ";"
If KeyCode = vbKeyT And edit = True Then edit = False: PlayMIDI: Editform.Hide: Movetimer.Enabled = True: Bltbas.Blt
If KeyCode = vbKeyS And edit = True Then savemap
If KeyCode = vbKeyD And edit = False Then GameForm.Height = 10050: Label1.Visible = True
If KeyCode = vbKeyUp And edit = False Then Keymov(0) = True
If KeyCode = vbKeyDown And edit = False Then Keymov(2) = True
If KeyCode = vbKeyLeft And edit = False Then Keymov(3) = True
If KeyCode = vbKeyRight And edit = False Then Keymov(1) = True
End Sub

Private Sub DXPic_KeyUp(KeyCode As Integer, Shift As Integer)
If Imgy < anim(walkinganim + 1) Then Imgy = 0
If KeyCode = vbKeyUp And edit = False Then Keymov(0) = False: Bltbas.Blt
If KeyCode = vbKeyDown And edit = False Then Keymov(2) = False: Bltbas.Blt
If KeyCode = vbKeyLeft And edit = False Then Keymov(3) = False: Bltbas.Blt
If KeyCode = vbKeyRight And edit = False Then Keymov(1) = False: Bltbas.Blt

End Sub

Private Sub DXPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = CInt(X / 16)
j = CInt(Y / 16)
For i = 1 To guinumber
If GUI(i, 0) = "MenuButton" Then
If Button = 1 Then
If X > GUI(i, 6) And X < GUI(i, 8) And Y > GUI(i, 7) And Y < GUI(i, 9) Then
quitmenu = Not quitmenu: Bltbas.Blt
End If
End If
End If
Next i
If quitmenu = True And X > 30 And X < 118 And Y > 58 And Y < 72 And Button = 1 Then End
If quitmenu = True And X > 30 And X < 67 And Y > 77 And Y < 91 And Button = 1 Then End
End Sub

Private Sub DXPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If edit = False Then
For i = 1 To guinumber
If GUI(i, 10) <> "" Then
If X > GUI(i, 6) And X < GUI(i, 8) And Y > GUI(i, 7) And Y < GUI(i, 9) Then
If GUI(i, 11) = True Then
Bltbas.bltmouseover (i)
ElseIf Left(GUI(i, 0), 4) = "Menu" And quitmenu = True Then
Bltbas.bltmouseover (i)
End If
Else
If mouseov = i Then Bltbas.bltmouseout (i)
End If
End If
Next i
End If
End Sub

Private Sub DXPic_Paint()
If frmLoading.Visible = True Then frmLoading.Visible = False
Bltbas.Blt
End Sub

Private Sub Form_GotFocus()
If frmLogin.Visible = True Then frmLogin.SetFocus
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
frmLoading.Show
Playerdetails(8) = Splash.Text1(6).Text
Playerdetails(9) = Splash.Text1(5).Text
Playerdetails(10) = Splash.Text1(4).Text
Playerdetails(11) = Splash.Text1(3).Text
Playerdetails(12) = Splash.Text1(2).Text
Playerdetails(13) = Splash.Text1(1).Text
Playerdetails(14) = Splash.Text1(0).Text
Playerdetails(15) = Splash.Text1(7).Text
Getwindowcolours
LoadSettings
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
DXPic.Width = 640
DXPic.Height = 480
GameForm.Width = 9735
GameForm.Height = 7935
loadgui
loadmap ("big")
Me.Hide
Init
Me.Hide
Bltbas.bltmap
loadcharacter
If frmOptions.Check1.Value = 1 Then LoadMIDI App.Path & "\music.mid"
If frmOptions.Check2.Value = 1 Then LoadWave 1, App.Path & "\tsteps.wav"
If frmOptions.Check1.Value = 1 Then PlayMIDI
loadcharacterdetails
Stam = False
Playerdetails(2) = 10
Playerdetails(3) = 0
Playerdetails(4) = 10
Playerdetails(5) = 120
Playerdetails(6) = 13
Playerdetails(7) = 100
Me.Show
Imgx = 0
Imgy = 0
plx = 10
ply = 7.5
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set DX = Nothing
Set DD = Nothing
Set DS = Nothing
Set DSSound(100) = Nothing
Set DDMLoad = Nothing
Set DDMPerf = Nothing
Set DDMSeg = Nothing
Set DDPrimSurf = Nothing
Set DDMap = Nothing
Set DDHead = Nothing
Set DDCharacter = Nothing
Set DDGetReady = Nothing
Set DDTop = Nothing
Set DDBackround = Nothing
Set DDWeapons = Nothing
End
End Sub



Private Sub Movetimer_Timer()
'If Playerdetails(7) <= 0 Then MsgBox "DEAD": End
ply = Round(ply, 1)
plx = Round(plx, 1)
If Keymov(0) = False And Keymov(1) = False And Keymov(2) = False And Keymov(3) = False Then Exit Sub
If Keymov(2) = True Then
Collisiondetect ("down")
Imgx = 2
If coll = True Then GoTo 1
ply = ply + 0.2
End If
If Keymov(0) = True Then
Collisiondetect ("up")
Imgx = 0
If coll = True Then GoTo 1
ply = ply - 0.2
End If
If Keymov(1) = True Then
Collisiondetect ("right")
Imgx = 3
If coll = True Then GoTo 1
plx = plx + 0.2
End If
If Keymov(3) = True Then
Collisiondetect ("left")
Imgx = 1
If coll = True Then GoTo 1
plx = plx - 0.2
End If
Imgy = Imgy + 1
If Stam = True Then Playerdetails(7) = Playerdetails(7) - 1
1 detectground
If sit = True Then Imgy = anim(0)
If Imgy > walkinganim Then Imgy = anim(1): If frmOptions.Check2.Value = 1 Then DSSound(1).Play flag
If sit = True Then Imgy = anim(walkinganim + 1)
If coll = True Then Imgy = anim(0)
Bltbas.Blt
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text1.Text = "debugmode" Then MsgBox plx & ":" & ply: Exit Sub
If Left(Text1.Text, 7) = "stamoff" Then Let Stam = False
If Left(Text1.Text, 6) = "stamon" Then Let Stam = True
If Left(Text1.Text, 2) = "h=" Then
Playerdetails(2) = Right(Text1.Text, Len(Text1.Text) - 2)
Playerdetails(1) = "Health =" & Playerdetails(2)
Bltbas.Blt
Exit Sub
End If
If Left(Text1.Text, 3) = "sit" Then Imgy = 6: Blt
If Left(Text1.Text, 2) = "m=" Then
Playerdetails(3) = Right(Text1.Text, Len(Text1.Text) - 2)
Playerdetails(1) = "Money =" & Playerdetails(3)
Bltbas.Blt
Exit Sub
End If
If Left(Text1.Text, 3) = "lh=" Then
Playerdetails(4) = Right(Text1.Text, Len(Text1.Text) - 3)
Playerdetails(1) = "Left Hand Ammo =" & Playerdetails(4)
Bltbas.Blt
Exit Sub
End If
If Left(Text1.Text, 3) = "rh=" Then
Playerdetails(5) = Right(Text1.Text, Len(Text1.Text) - 3)
Playerdetails(1) = "Right Hand Ammo=" & Playerdetails(5)
Bltbas.Blt
Exit Sub
End If
If Left(Text1.Text, 2) = "b=" Then
Playerdetails(6) = Right(Text1.Text, Len(Text1.Text) - 2)
Playerdetails(1) = "Bag =" & Playerdetails(6)
Bltbas.Blt
Exit Sub
End If
If connected = True Then
Winsock1.SendData "spk:" & Text1.Text & ";"
End If
Playerdetails(1) = Text1.Text
Text1.Text = ""
Bltbas.Blt
End If
End Sub

Private Sub Timer1_Timer()
Winsock1.SendData "charpos:" & plx & "," & ply & "," & Imgx & "," & Imgy
End Sub


Private Sub Timer2_Timer()
'For i = 1 To NPCno
'If left(NPC(0,NPCno),7)= "iftouch" and
'Next i
End Sub

Private Sub Winsock1_Connect()
Dim logintext As String
logintext = "login:" & frmLogin.txtUserName.Text & ":" & frmLogin.txtPassword.Text
Winsock1.SendData logintext
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim incomdata As String
Winsock1.GetData incomdata
DataProcess incomdata
End Sub

