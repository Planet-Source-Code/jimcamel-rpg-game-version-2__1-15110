Attribute VB_Name = "Declares"
Public walkinganim
Public anim(50)
Public NPCno
Public NPC(50, 500)
Public mapplx
Public mapply
Public Stam As Boolean
Public sit As Boolean
Public characterdata(50, 10, 4, 6) As Integer
Public characterdataname(50)
Public feetrect As RECT
Public mouseov As Integer
Public plonline
Public connected As Boolean
Public oppdetail(50, 25)
Public transcoltemp
Public walkstartx
Public walkstarty
Public Dir
Public walkx1
Public walky1
Public edit As Boolean
Public Keymov(4) As Boolean
Public coll As Boolean
Public amenu As Boolean
Public smenu As Boolean
Public dmenu As Boolean
Public quitmenu As Boolean
Public DDfont As New StdFont
Public bInit As Boolean
Public Tempfile As String
Public kPress(5) As Boolean
Public r1 As RECT
Public r2 As RECT
Public Imgx As Integer
Public Imgy As Integer
Public plx
Public ply
Public map(300, 300, 3)
Public maplink(50, 7)
Public ddmapsize As RECT
Public linkno As Integer

Public noguisurface As Integer
Public guinumber As Integer
Public GUI(50, 20)
Public DDGUI(50) As DirectDrawSurface7
Public DX As New DirectX7
Public DD As DirectDraw7
Public DS As DirectSound
Public DSSound(100) As DirectSoundBuffer
Public DDMLoad As DirectMusicLoader
Public DDMPerf As DirectMusicPerformance
Public DDMSeg As DirectMusicSegment
Public DDPrimSurf As DirectDrawSurface7
Public DDOppbuffer As DirectDrawSurface7
Public DDOpphead(50) As DirectDrawSurface7
Public DDOpponents(50) As DirectDrawSurface7
Public DDMap As DirectDrawSurface7
Public DDHead As DirectDrawSurface7
Public DDmapbuffer As DirectDrawSurface7
Public DDGrass As DirectDrawSurface7
Public DDFeet As DirectDrawSurface7
Public DDTorso As DirectDrawSurface7
Public DDHands As DirectDrawSurface7
Public DDlegs As DirectDrawSurface7
Public DDarms As DirectDrawSurface7
Public DDbelt As DirectDrawSurface7
Public DDCharacter As DirectDrawSurface7
Public DDGetReady As DirectDrawSurface7
Public DDTop As DirectDrawSurface7
Public DDBackround As DirectDrawSurface7
Public DDWeapons As DirectDrawSurface7
Public ddClipper As DirectDrawClipper
Public DDNPC As DirectDrawSurface7
Public DDNPCtemp(50) As DirectDrawSurface7

Public RedShiftLeft As Long
Public RedShiftRight As Long
Public GreenShiftLeft As Long
Public GreenShiftRight As Long
Public BlueShiftLeft As Long
Public BlueShiftRight As Long
Private hdesktopwnd
Private hdccaps
Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Global colourdisplay As Integer
Global soundinit As Boolean
Global Is16bitc As Boolean
Global musicinit As Boolean
Global Playerdetails(20)
'playerdetails(0) = nickname
'playerdetails(1) = saying
'playerdetails(2) = health
'playerdetails(3) = money
'playerdetails(4) = left
'playerdetails(5) = right
'playerdetails(6) = bag
'playerdetails(7) = stamina
'playerdetails(8) = feet
'playerdetails(9) = legs
'playerdetails(10) = torso
'playerdetails(11) = belt
'playerdetails(12) = arms
'playerdetails(13) = hands
'playerdetails(14) = head
'playerdetails(15) = map
Dim DownloadedText As String, Downloading As Boolean
Global startx: Global starty
Global X1: Global Y1
Global rWin As RECT
Private Const SRCCOPY = &HCC0020
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Function CreateSurfaceFromFile(DirectDraw As DirectDraw7, ByVal Filename As String, SurfaceDesc As DDSURFACEDESC2) As DirectDrawSurface7
    Dim Picture As StdPicture
    Dim Width As Long
    Dim Height As Long
    Dim Surface As DirectDrawSurface7
    Dim hdcPicture As Long
    Dim hdcSurface As Long
    'If filetype = "GIF" Then
    Set Picture = LoadPicture(Filename)
    
    Width = CLng((Picture.Width * 0.001) * 567 / Screen.TwipsPerPixelX)
    Height = CLng((Picture.Height * 0.001) * 567 / Screen.TwipsPerPixelY)
    With SurfaceDesc
        If .lFlags = 0 Then .lFlags = DDSD_CAPS
        .lFlags = .lFlags Or DDSD_WIDTH Or DDSD_HEIGHT
        If .ddsCaps.lCaps = 0 Then .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        If .lWidth = 0 Then .lWidth = Width
        If .lHeight = 0 Then .lHeight = Height
    End With
    Set Surface = DirectDraw.CreateSurface(SurfaceDesc)
    hdcPicture = CreateCompatibleDC(0)
    SelectObject hdcPicture, Picture.Handle
        hdcSurface = Surface.GetDC
    StretchBlt hdcSurface, 0, 0, SurfaceDesc.lWidth, SurfaceDesc.lHeight, hdcPicture, 0, 0, Width, Height, SRCCOPY
    Surface.ReleaseDC hdcSurface
    
    'GetGifInfo (Filename)
    'If GifInfo.Transparent = True Then
    Dim ddtrans1 As DDCOLORKEY
        If colourdisplay = 16 Then
    coltran = DDColor(RGB(255, 0, 255))
    Else
    coltran = RGB(255, 0, 255)
    End If
    ddtrans1.low = coltran
    ddtrans1.high = coltran
    Surface.SetColorKey DDCKEY_SRCBLT, ddtrans1
    
    'ElseIf filetype = "PNG" Then
    'LoadPNG Filename, Img
    'Width = Img.ImageWidth
    'Height = Img.ImageHeight
    'With SurfaceDesc
    '    If .lFlags = 0 Then .lFlags = DDSD_CAPS
    '    .lFlags = .lFlags Or DDSD_WIDTH Or DDSD_HEIGHT
    '    If .ddsCaps.lCaps = 0 Then .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    '    If .lWidth = 0 Then .lWidth = Width
    '    If .lHeight = 0 Then .lHeight = Height
    'End With
    'Set Surface = DirectDraw.CreateSurface(SurfaceDesc)
    'hdcPicture = CreateCompatibleDC(0)
    'SelectObject hdcPicture, Img.Handle
    'hdcSurface = Surface.GetDC
    'StretchBlt hdcSurface, 0, 0, SurfaceDesc.lWidth, SurfaceDesc.lHeight, hdcPicture, 0, 0, Width, Height, SRCCOPY
    'Surface.ReleaseDC hdcSurface
    'ElseIf filetype = "" Then Exit Function
    'End If
    
        DeleteDC hdcPicture
    Set Picture = Nothing
    Set CreateSurfaceFromFile = Surface
    Set Surface = Nothing
End Function
Public Sub Getwindowcolours()
Dim DisplayBits
Dim RetVal
hdccaps = GetDC(hdesktopwnd)
DisplayBits = GetDeviceCaps(hdccaps, 12)
RetVal = ReleaseDC(hdesktopwnd, hdccaps)
colourdisplay = DisplayBits
If colourdisplay = 24 Then Let colourdisplay = 32
If colourdisplay <> 16 And colourdisplay <> 32 Then MsgBox "You must have 16/24 or 32 bit colour to run this program"
End Sub
Public Function DDRGB(red As Long, green As Long, blue As Long) As Long
DDRGB = (red \ RedShiftRight) * RedShiftLeft + (green \ GreenShiftRight) * GreenShiftLeft + (blue \ BlueShiftRight) * BlueShiftLeft
End Function

Public Sub GetRGBfromColor(ByVal Color As Long, ByRef red As Long, ByRef green As Long, ByRef blue As Long)
Dim HexadecimalValue As String
    HexadecimalValue = Hex(Val(Color))
    
    If Len(HexadecimalValue) < 6 Then
        HexadecimalValue = String(6 - Len(HexadecimalValue), "0") + HexadecimalValue
    End If
    blue = CLng("&H" + Mid(HexadecimalValue, 1, 2))
    green = CLng("&H" + Mid(HexadecimalValue, 3, 2))
    red = CLng("&H" + Mid(HexadecimalValue, 5, 2))
End Sub


Public Function DDColor(RGBColor As Long) As Long
    Dim RedVal As Long
    Dim GreenVal As Long
    Dim BlueVal As Long
    GetRGBfromColor RGBColor, RedVal, GreenVal, BlueVal
    DDColor = DDRGB(RedVal, GreenVal, BlueVal)
    End Function

Public Function Loaddirectxsurface(DirectDrawSurface As DirectDrawSurface7, ByVal Filename As String, Optional Width As Long, Optional Height As Long) As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2
'Checkfiletype (Filename)
'If filetype = "GIF" Then
    GameForm.OpenPict.Picture = LoadPicture(Filename)
'    ElseIf filetype = "PNG" Then
'    LoadPNG Filename, Img
'    GameForm.OpenPict.Width = Img.ImageWidth + (GameForm.OpenPict.Width - GameForm.OpenPict.ScaleWidth)
'    GameForm.OpenPict.Height = Img.ImageHeight + (GameForm.OpenPict.Height - GameForm.OpenPict.ScaleHeight)
'    DrawImage GameForm.OpenPict.hdc, Img
'    ElseIf filetype = "" Then Exit Function
'    End If
If Width <> 0 Then
    Tempfile = Filename
    ddsd1.lFlags = DDSD_CAPS
    ddsd1.lHeight = Height
    ddsd1.lWidth = Width
    Set DirectDrawSurface = CreateSurfaceFromFile(DD, Tempfile, ddsd1)
    Else
    Tempfile = Filename
'    GameForm.OpenPict.Picture = LoadPicture(Tempfile)
    ddsd1.lFlags = DDSD_CAPS
    ddsd1.lHeight = GameForm.OpenPict.Height
    ddsd1.lWidth = GameForm.OpenPict.Width
    Set DirectDrawSurface = CreateSurfaceFromFile(DD, Tempfile, ddsd1)
End If
    DoEvents
End Function
