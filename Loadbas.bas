Attribute VB_Name = "Work"
Sub loadcharacterdetails()
Dim temps(50)
Dim tempno
Dim sittemp
Dim charnum As Integer
Dim linenumber As Integer
filenum = FreeFile
Dim filedata As String
Open Playerdetails(15) For Input As #filenum
Do
Line Input #filenum, filedata
If Left(filedata, 1) = "'" Or filedata = "" Then GoTo 1
If Left(filedata, 1) = "!" Then
Let charnum = Mid(filedata, 2, InStr(2, filedata, " ") - 1)
Let characterdataname(charnum) = Right(filedata, Len(filedata) - InStr(InStr(1, filedata, "="), filedata, " "))
If characterdataname(charnum) = "Idle" Then
anim(0) = charnum
ElseIf Left(characterdataname(charnum), 7) = "Walking" Then
anim(Right(characterdataname(charnum), Len(characterdataname(charnum)) - 8)) = charnum: walkinganim = charnum
ElseIf characterdataname(charnum) = "Sitting" Then sittemp = charnum
Else
temps(tempno) = charnum
tempno = tempno + 1
End If
Let linenumber = 0
i = 1
a = 1
Else
Do
Direction = Mid(filedata, InStr(1, filedata, ":") + 1, 1)
temp = Right(filedata, Len(filedata) - InStr(1, filedata, ";"))
If InStr(1, temp, " ") > 0 Then
filedata = Left(temp, InStr(1, temp, " ") - 1)
Else
filedata = Left(temp, Len(temp))
End If
i = 1
a = 1
Do
characterdata(charnum, linenumber, Direction, a) = Mid(filedata, i, InStr(i, filedata, ";") - i)
Let i = InStr(i, filedata, ";") + 1
a = a + 1
If a = 6 Then
If Right(filedata, 1) = " " Then
Let characterdata(charnum, linenumber, Direction, 6) = Right(filedata, InStr(i, filedata, " "))
Else
Let characterdata(charnum, linenumber, Direction, 6) = Right(filedata, Len(filedata) - (i - 1))
End If
End If
Loop While a < 6
For j = 1 To 6
Next j
filedata = Right(temp, Len(temp) - InStr(1, temp, ":") + 1)
Loop While InStr(1, temp, ":") > 0
End If
Let linenumber = linenumber + 1
1 Loop While Not EOF(1)
Close #filenum
anim(walkinganim + 1) = sittemp
If tempno = 0 Then Exit Sub
For i = 0 To tempno
anim(walkinganim + 2 + i) = temps(i)
Next i
End Sub


Sub detectground()
On Error Resume Next
sit = False
If plx > 10 Then
plx2 = CInt(plx + mapplx) / 2
Else
plx2 = CInt(plx)
End If
coll = False
If ply > 7.5 Then
ply2 = CInt(ply + mapply) / 2
Else
ply2 = CInt(ply)
End If
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
a = (plx2 * 2) + 1
b = (ply2 * 2) + 2
For j = 1 To linkno
c = maplink(j, 1)
d = maplink(j, 3)
e = maplink(j, 2)
f = maplink(j, 4)
'If j = 3 Then MsgBox a & ":" & b & ":" & c & ":" & d & ":" & e & ":" & f
If a >= c Then
If a <= d Then
If b >= e Then
If b <= f Then
If maplink(j, 6) <> "" Then Let plx = (maplink(j, 6))
If maplink(j, 7) <> "" Then Let ply = (maplink(j, 7))
loadmap (maplink(j, 5))
Bltbas.bltmap
Exit Sub
End If
End If
End If
End If
Next j
For i = 0 To 1
If map((plx2 * 2) + i, (ply2 * 2) + 2, 3) = 3 Then sit = True
Next i
End Sub
Sub loadgui()
Dim filedata As String
Dim temp As String
Dim data As Integer
filenum = FreeFile
Open Splash.Text1(8).Text For Input As #filenum
Do
Line Input #filenum, filedata
If Left(filedata, 1) = "'" Then GoTo 1

If Left(filedata, 4) = "Data" Then
    data = -1
    filedata = Right(filedata, Len(filedata) - 5)
    Do
        Let temp = Right(filedata, Len(filedata) - InStr(1, filedata, ";"))
        filedata = Left(filedata, InStr(1, filedata, ";") - 1)
        If data = -1 Then Let guinumber = filedata
        If data > -1 Then GUI(guinumber, data) = filedata
        filedata = temp
        data = data + 1
    Loop While InStr(1, filedata, ";") > 0
End If

If Left(filedata, 3) = "SRC" Then
    data = 2
    filedata = Right(filedata, Len(filedata) - 4)
    Do
        Let temp = Right(filedata, Len(filedata) - InStr(1, filedata, ";"))
        filedata = Left(filedata, InStr(1, filedata, ";") - 1)
        GUI(guinumber, data) = filedata
        filedata = temp
        data = data + 1
    Loop While InStr(1, filedata, ";") > 0
End If
If Left(filedata, 4) = "DEST" Then
    data = 6
    filedata = Right(filedata, Len(filedata) - 5)
    Do
        Let temp = Right(filedata, Len(filedata) - InStr(1, filedata, ";"))
        filedata = Left(filedata, InStr(1, filedata, ";") - 1)
        GUI(guinumber, data) = filedata
        filedata = temp
        data = data + 1
    Loop While InStr(1, filedata, ";") > 0
End If

If Left(filedata, 9) = "Mouseover" Then
    data = 10
    filedata = Right(filedata, Len(filedata) - 10)
    Do
        Let temp = Right(filedata, Len(filedata) - InStr(1, filedata, ";"))
        filedata = Left(filedata, InStr(1, filedata, ";") - 1)
        GUI(guinumber, data) = filedata
        filedata = temp
        data = data + 1
    Loop While InStr(1, filedata, ";") > 0
End If

If Left(filedata, 7) = "Visible" Then
    data = 11
    filedata = Right(filedata, Len(filedata) - 8)
    Do
        Let temp = Right(filedata, Len(filedata) - InStr(1, filedata, ";"))
        filedata = Left(filedata, InStr(1, filedata, ";") - 1)
        GUI(guinumber, data) = filedata
        filedata = temp
        data = data + 1
    Loop While InStr(1, filedata, ";") > 0
End If

If Left(filedata, 4) = "Rows" Then
    data = 13
    filedata = Right(filedata, Len(filedata) - 5)
    Do
        Let temp = Right(filedata, Len(filedata) - InStr(1, filedata, ";"))
        filedata = Left(filedata, InStr(1, filedata, ";") - 1)
        GUI(guinumber, data) = filedata
        filedata = temp
        data = data + 1
    Loop While InStr(1, filedata, ";") > 0
End If

1 Loop While Not EOF(1)
Close #filenum
countloadsurfaces
End Sub
Sub countloadsurfaces()
noguisurface = 1
GUI(1, 12) = 1
For i = 2 To guinumber
For j = 1 To i
If GUI(j, 1) = GUI(i, 1) Then Let GUI(i, 12) = GUI(j, 12): Exit For
Next j
If GUI(i, 12) = "" Then noguisurface = noguisurface + 1: Let GUI(i, 12) = noguisurface
Next i
End Sub
Sub loadcharacter()
DDCharacter.Blt feetrect, DDFeet, feetrect, DDBLT_WAIT
DDCharacter.Blt feetrect, DDlegs, feetrect, DDBLT_KEYSRC
DDCharacter.Blt feetrect, DDTorso, feetrect, DDBLT_KEYSRC
DDCharacter.Blt feetrect, DDbelt, feetrect, DDBLT_KEYSRC
DDCharacter.Blt feetrect, DDHands, feetrect, DDBLT_KEYSRC
DDCharacter.Blt feetrect, DDarms, feetrect, DDBLT_KEYSRC
Set DDFeet = Nothing
Set DDlegs = Nothing
Set DDTorso = Nothing
Set DDbelt = Nothing
Set DDHands = Nothing
Set DDarms = Nothing
End Sub
Public Sub MaskToShiftValues(ByVal Mask As Long, ShiftRight As Long, ShiftLeft As Long)
    Dim ZeroBitCount As Long
    Dim OneBitCount As Long
    ZeroBitCount = 0
    Do While (Mask And 1) = 0
    ZeroBitCount = ZeroBitCount + 1
    Mask = Mask \ 2
    Loop
    OneBitCount = 0
    Do While (Mask And 1) = 1
    OneBitCount = OneBitCount + 1
    Mask = Mask \ 2
    Loop
    ShiftRight = 2 ^ (8 - OneBitCount)
    ShiftLeft = 2 ^ ZeroBitCount
    End Sub

Sub Stopwave()
If soundinit = False Then Exit Sub
DSSound(1).Stop
DSSound(1).SetCurrentPosition 0
End Sub
Sub LoadWave(i As Integer, sfile As String)
    Dim bufferDesc As DSBUFFERDESC
    Dim DSstepformat As WAVEFORMATEX
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    DSstepformat.nFormatTag = WAVE_FORMAT_PCM
    DSstepformat.nChannels = 2
    DSstepformat.lSamplesPerSec = 22050
    DSstepformat.nBitsPerSample = 16
    DSstepformat.nBlockAlign = DSstepformat.nBitsPerSample / 8 * DSstepformat.nChannels
    DSstepformat.lAvgBytesPerSec = DSstepformat.lSamplesPerSec * DSstepformat.nBlockAlign
    Set DSSound(i) = DS.CreateSoundBufferFromFile(sfile, bufferDesc, DSstepformat)
End Sub

Sub initsound()
    Set DS = DX.DirectSoundCreate("")
    DS.SetCooperativeLevel GameForm.hwnd, DSSCL_PRIORITY
    If Err.Number <> DD_OK Then MsgBox "Error initializing DirectSound", vbExclamation, "Error"
soundinit = True
End Sub
Sub LoadMIDI(Filename As String)
On Error Resume Next
Set DDMSeg = DDMLoad.LoadSegment(Filename)
If Err.Number <> DD_OK Then MsgBox "ERROR : Could not load MIDI file!", vbExclamation, "ERROR!"
End Sub

Sub PlayMIDI()
On Error Resume Next
DDMPerf.SetMasterVolume 5
DDMPerf.PlaySegment DDMSeg, 0, 0
End Sub

Sub StopMIDI()
On Error Resume Next
DDMPerf.Stop DDMSeg, Nothing, 0, 0
End Sub

Sub iniMIDI()
On Error Resume Next
Set DDMLoad = DX.DirectMusicLoaderCreate
Set DDMPerf = DX.DirectMusicPerformanceCreate
DDMPerf.Init Nothing, hwnd
DDMPerf.SetPort -1, 1
DDMPerf.SetMasterAutoDownload True
If Err.Number <> DD_OK Then MsgBox "Error initializing DirectMusic", vbExclamation, "Error"
musicinit = True
End Sub
Sub savemap()
filenum = FreeFile
Open App.Path & "/maps/map.txt" For Output As #filenum
For i = 0 To 50
For j = 0 To 50
Print #filenum, map(j, i, 1) & map(j, i, 2) & map(j, i, 3);
Next j
Print #filenum, ","
Next i
Close #filenum
MsgBox "Map Saved"
End Sub
Sub loadmap(mapfile As String)
linkno = 0
NPCno = 0
Dim Datamap As String
If Left(mapfile, 5) <> "/maps" Then Let mapfile = "/maps/" & mapfile
If Right(mapfile, 4) <> ".map" Then Let mapfile = mapfile & ".map"
filenum = FreeFile
Open App.Path & mapfile For Input As #filenum
Let j = 0
Let k = 0
Do
Input #filenum, Datamap
If Datamap = "" Then GoTo 1
If Left(Datamap, 5) = "*link" Then GoTo Link
For i = 1 To Len(Datamap) Step 3
map(j, k, 1) = Asc(Mid(Datamap, i, 1)) - 48: map(j, k, 2) = Asc(Mid(Datamap, i + 1, 1)) - 48: map(j, k, 3) = Asc(Mid(Datamap, i + 2, 1)) - 48
j = j + 1
Next i
ddmapsize.Right = j * 32
j = 0
k = k + 1
1 Loop While Not EOF(1)
Close #filenum
ddmapsize.Left = 0: ddmapsize.Top = 0: ddmapsize.Bottom = k * 32
Exit Sub
Link:
linkno = linkno + 1
a = 7
b = InStr(a, Datamap, ";")
c = InStr(b + 1, Datamap, ";")
d = InStr(c + 1, Datamap, ";")
e = InStr(d + 1, Datamap, ";")
If InStr(e + 1, Datamap, ";") > 0 Then
f = InStr(e + 1, Datamap, ";")
maplink(linkno, 5) = Mid(Datamap, e + 1, f - (e + 1))
If InStr(f + 1, Datamap, ";") > 0 Then
g = InStr(f + 1, Datamap, ";")
maplink(linkno, 6) = Mid(Datamap, f + 1, g - (f + 1))
maplink(linkno, 7) = Right(Datamap, Len(Datamap) - g)
Else
maplink(linkno, 6) = Right(Datamap, Len(Datamap) - f)
End If
Else
maplink(linkno, 5) = Right(Datamap, Len(Datamap) - e)
End If
maplink(linkno, 1) = Mid(Datamap, a, b - a)
maplink(linkno, 2) = Mid(Datamap, b + 1, c - (b + 1))
maplink(linkno, 3) = Mid(Datamap, c + 1, d - (c + 1))
maplink(linkno, 4) = Mid(Datamap, d + 1, e - (d + 1))
GoTo 1
End Sub
Sub Collisiondetect(Dir As String)
On Error Resume Next
coll = False
If plx > 10 Then
plx2 = CInt(plx + mapplx) / 2
Else
plx2 = CInt(plx)
End If
coll = False
If ply > 7.5 Then
ply2 = CInt(ply + mapply) / 2
Else
ply2 = CInt(ply)
End If
Select Case Dir
Case "up"
For i = 0 To 1
If map((plx2 * 2) + i, (ply2 * 2), 3) = 1 Then coll = True
Next i
Case "down"
For i = 0 To 1
'If map((plx2 * 2) + i, (ply2 * 2) + 2, 3) = 1 Then coll = True
Next i
Case "left"
For i = 1 To 2
If map((plx2 * 2), (ply2 * 2) + i, 3) = 1 Then coll = True
Next i
Case "right"
For i = 1 To 2
If map((plx2 * 2) + 2, (ply2 * 2) + i, 3) = 1 Then coll = True
Next i
End Select
End Sub
Sub DataProcess(datastr As String)
Dim temp As String
Dim redo As Boolean
1
If Len(datastr) > InStr(1, datastr, ";") Then
Let temp = Right(datastr, Len(datastr) - InStr(1, datastr, ";"))
datastr = Left(datastr, InStr(1, datastr, ";"))
redo = True
If Len(datastr) = 0 Then Exit Sub
Else
redo = False
End If
GameForm.Label1.Caption = datastr
If Left(datastr, 4) = "plc:" Then
If connected <> True Then
    Unload frmLogin
    GameForm.SetFocus
    connected = True
    GameForm.Timer1.Enabled = True
End If
plonline = Mid(datastr, 5, InStr(5, datastr, ";") - 5)
frmplayerlist.Frame1.Caption = "Players Online: plonline"
Let GameForm.Label2.Caption = plonline
End If
If InStr(1, datastr, "you:") <> 0 Then
bob = Mid(datastr, InStr(1, datastr, "you:"), InStr(InStr(1, datastr, "you:"), datastr, ";") - InStr(1, datastr, "you:"))
Do
bob = Right(bob, Len(bob) - InStr(1, bob, ":"))
Loop While InStr(1, bob, ":") <> 0
a = InStr(InStr(1, bob, ":") + 1, bob, ",")
b = InStr(a + 1, bob, ",")
c = InStr(b + 1, bob, ",")
d = InStr(c + 1, bob, ",")
e = InStr(d + 1, bob, ",")
f = InStr(e + 1, bob, ",")
g = InStr(f + 1, bob, ",")
h = InStr(g + 1, bob, ",")
i = InStr(h + 1, bob, ",")
j = InStr(i + 1, bob, ",")
k = InStr(j + 1, bob, ",")
l = InStr(k + 1, bob, ",")
m = InStr(l + 1, bob, ",")
n = InStr(m + 1, bob, ",")
plx = Mid(bob, 1, a - 1)
ply = Mid(bob, a + 1, b - (a + 1))
imx = Mid(bob, b + 1, c - (b + 1))
imy = Mid(bob, c + 1, d - (c + 1))
Playerdetails(0) = Mid(bob, d + 1, e - (d + 1))
loadmap (Mid(bob, e + 1, f - (e + 1)))
Playerdetails(2) = Mid(bob, f + 1, g - (f + 1))
Playerdetails(8) = App.Path & "\graphics\" & Mid(bob, g + 1, h - (g + 1))
Playerdetails(9) = App.Path & "\graphics\" & Mid(bob, h + 1, i - (h + 1))
Playerdetails(10) = App.Path & "\graphics\" & Mid(bob, i + 1, j - (i + 1))
Playerdetails(11) = App.Path & "\graphics\" & Mid(bob, j + 1, k - (j + 1))
Playerdetails(12) = App.Path & "\graphics\" & Mid(bob, k + 1, l - (k + 1))
Playerdetails(13) = App.Path & "\graphics\" & Mid(bob, l + 1, m - (l + 1))
Playerdetails(14) = App.Path & "\graphics\" & Mid(bob, m + 1, n - (m + 1))
Playerdetails(15) = App.Path & "\graphics\" & Right(bob, Len(bob) - n) 'Mid(bob, n + 1, o - (n + 1))
Playerdetails(15) = Left(Playerdetails(15), Len(Playerdetails(15)) - 1)
Reinit
'Blt
End If

If datastr = "failwrongpass" Then MsgBox "Incorrect Password"
If datastr = "failnouser" Then MsgBox "No such user"
If Left(datastr, 14) = "ServerMessage:" Then MsgBox Right(datastr, Len(datastr) - 14), vbExclamation, "Server Message"
If Left(datastr, 4) = "1plr" Then
MsgBox datastr
bob = Left(datastr, InStr(1, datastr, ";") - 1)
bob = Right(bob, Len(bob) - InStr(6, bob, ":"))
a = InStr(1, bob, ",")
b = InStr(a + 1, bob, ",")
c = InStr(b + 1, bob, ",")
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 1) = Mid(bob, 1, a - 1)
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 2) = Mid(bob, a + 1, b - (a + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 3) = Mid(bob, b + 1, c - (b + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 4) = Right(bob, Len(bob) - c)
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 4) = Left(oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 4), Len(oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 4)) - 1)
Bltbas.Blt
End If

If Left(datastr, 3) = "plr" Then
'MsgBox datastr
bob = Left(datastr, InStr(1, datastr, ";") - 1)
bob = Right(bob, Len(bob) - InStr(5, bob, ":"))
a = InStr(1, bob, ",")
b = InStr(a + 1, bob, ",")
c = InStr(b + 1, bob, ",")
d = InStr(c + 1, bob, ",")
e = InStr(d + 1, bob, ",")
f = InStr(e + 1, bob, ",")
g = InStr(f + 1, bob, ",")
h = InStr(g + 1, bob, ",")
i = InStr(h + 1, bob, ",")
j = InStr(i + 1, bob, ",")
k = InStr(j + 1, bob, ",")
l = InStr(k + 1, bob, ",")
m = InStr(l + 1, bob, ",")
n = InStr(m + 1, bob, ",")
'oppdetail
'0 = name
'1 = x
'2 = y
'3 = dir
'4 = fram
'5 = map
'6 = health
'7 = feet
'8 = legs
'9 = torso
'10 = belt
'11 = arms
'12 = hands
'13 = head
'14 = map
'15 = speak
If Mid(bob, g + 1, h - (g + 1)) = "" Then Exit Sub
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 1) = Mid(bob, 1, a - 1)
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 2) = Mid(bob, a + 1, b - (a + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 3) = Mid(bob, b + 1, c - (b + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 4) = Mid(bob, c + 1, d - (c + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 0) = Mid(bob, d + 1, e - (d + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 5) = Mid(bob, e + 1, f - (e + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 6) = Mid(bob, f + 1, g - (f + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 7) = App.Path & "\graphics\" & Mid(bob, g + 1, h - (g + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 8) = App.Path & "\graphics\" & Mid(bob, h + 1, i - (h + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 9) = App.Path & "\graphics\" & Mid(bob, i + 1, j - (i + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 10) = App.Path & "\graphics\" & Mid(bob, j + 1, k - (j + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 11) = App.Path & "\graphics\" & Mid(bob, k + 1, l - (k + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 12) = App.Path & "\graphics\" & Mid(bob, l + 1, m - (l + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 13) = App.Path & "\graphics\" & Mid(bob, m + 1, n - (m + 1))
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 14) = App.Path & "\graphics\" & Right(bob, Len(bob) - n)
oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 14) = Left(oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 14), Len(oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 14)) - 1)
frmplayerlist.List1.AddItem (oppdetail(Mid(datastr, 5, InStr(5, datastr, ")") - 5), 0))
Loadopp (Mid(datastr, 5, InStr(5, datastr, ")") - 5))
Bltbas.Blt
End If
If redo = True Then Let datastr = temp: GoTo 1
End Sub
Sub Loadopp(no As Integer)
    feetrect.Left = 0: feetrect.Top = 0
'    GetGifInfo (oppdetail(no, 7))
    feetrect.Bottom = GifInfo.Height: feetrect.Right = GifInfo.Width
    Loaddirectxsurface DDFeet, oppdetail(no, 7)
    Loaddirectxsurface DDlegs, oppdetail(no, 8)
    Loaddirectxsurface DDTorso, oppdetail(no, 9)
    Loaddirectxsurface DDbelt, oppdetail(no, 10)
    Loaddirectxsurface DDarms, oppdetail(no, 11)
    Loaddirectxsurface DDHands, oppdetail(no, 12)
    Loaddirectxsurface DDOpponents(no), oppdetail(no, 7)
    Loaddirectxsurface DDOpphead(no), oppdetail(no, 13)
DDOpponents(no).Blt feetrect, DDFeet, feetrect, DDBLT_WAIT
DDOpponents(no).Blt feetrect, DDlegs, feetrect, DDBLT_KEYSRC
DDOpponents(no).Blt feetrect, DDTorso, feetrect, DDBLT_KEYSRC
DDOpponents(no).Blt feetrect, DDbelt, feetrect, DDBLT_KEYSRC
DDOpponents(no).Blt feetrect, DDHands, feetrect, DDBLT_KEYSRC
DDOpponents(no).Blt feetrect, DDarms, feetrect, DDBLT_KEYSRC
Set DDFeet = Nothing
Set DDlegs = Nothing
Set DDTorso = Nothing
Set DDbelt = Nothing
Set DDHands = Nothing
Set DDarms = Nothing
End Sub

Sub Init()
If frmOptions.Check1.Value = 1 Then iniMIDI
If frmOptions.Check2.Value = 1 Then initsound
    Set DD = DX.DirectDrawCreate("")
    Call DD.SetCooperativeLevel(GameForm.hwnd, DDSCL_NORMAL)
DDfont.Name = "Verdana"
DDfont.Size = 11
DDfont.Bold = True

    Dim ddsd2 As DDSURFACEDESC2
    'Visible surface
    ddsd2.lFlags = DDSD_CAPS
    ddsd2.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set DDPrimSurf = DD.CreateSurface(ddsd2)
    Dim PixelFormat As DDPIXELFORMAT
    DDPrimSurf.GetPixelFormat PixelFormat
    MaskToShiftValues PixelFormat.lRBitMask, RedShiftRight, RedShiftLeft
    MaskToShiftValues PixelFormat.lGBitMask, GreenShiftRight, GreenShiftLeft
    MaskToShiftValues PixelFormat.lBBitMask, BlueShiftRight, BlueShiftLeft
    
    ScaleMode = 3
    
    For i = 1 To noguisurface
    For j = 1 To guinumber
    If GUI(j, 12) = i Then
        DoEvents
        Loaddirectxsurface DDGUI(i), App.Path & GUI(j, 1)
    Exit For
    End If
    Next j
    Next i
    DoEvents
    UpdateLoading (5)
    'For i = 1 To NPCno
    ' If Left(NPC(0, NPCno), 3) = "img" Then
    ' name =
    ' Loaddirectxsurface DDGUI(i), App.Path & "/graphics/" GUI(j, 1)
    'Next i
        Loaddirectxsurface DDNPC, Playerdetails(13), 640, 480
    UpdateLoading (10)
    ' Character
    DoEvents
    Loaddirectxsurface DDmapbuffer, Playerdetails(13), 640, 480
    UpdateLoading (15)
    feetrect.Left = 0: feetrect.Top = 0
    GetGifInfo (Playerdetails(8))
    feetrect.Bottom = GifInfo.Height: feetrect.Right = GifInfo.Width
    Loaddirectxsurface DDFeet, Playerdetails(8)
    UpdateLoading (20)
    Loaddirectxsurface DDlegs, Playerdetails(9)
    UpdateLoading (25)
    Loaddirectxsurface DDTorso, Playerdetails(10)
    UpdateLoading (30)
    Loaddirectxsurface DDbelt, Playerdetails(11)
    UpdateLoading (35)
    Loaddirectxsurface DDarms, Playerdetails(12)
    UpdateLoading (40)
    Loaddirectxsurface DDHands, Playerdetails(13)
    UpdateLoading (45)
    Loaddirectxsurface DDCharacter, Playerdetails(8)
    UpdateLoading (50)
    Loaddirectxsurface DDHead, Playerdetails(14)
   'GUI
    UpdateLoading (55)
    Loaddirectxsurface DDMap, App.Path & "/graphics/nothing.gif", ddmapsize.Right, ddmapsize.Bottom
    UpdateLoading (60)
        Loaddirectxsurface DDOppbuffer, App.Path & "/graphics/nothing.gif", ddmapsize.Right, ddmapsize.Bottom
    UpdateLoading (65)
    Loaddirectxsurface DDWeapons, App.Path & "/graphics/nothing.gif"
    UpdateLoading (70)
    Loaddirectxsurface DDBackround, App.Path & "/graphics/grass.gif"
    UpdateLoading (75)
    Loaddirectxsurface DDTop, App.Path & "/graphics/top.gif"
   UpdateLoading (80)
     Loaddirectxsurface DDGetReady, App.Path & "/graphics/nothing.gif", 800, 600
    UpdateLoading (85)
    Loaddirectxsurface DDGrass, App.Path & "/graphics/nothing.gif"
    UpdateLoading (100)
    
   'Initilise Clipper
    Set ddClipper = DD.CreateClipper(0)
    ddClipper.SetHWnd GameForm.DXPic.hwnd
    DDPrimSurf.SetClipper ddClipper
    
    DDmapbuffer.SetFont DDfont
    DDmapbuffer.SetForeColor RGB(255, 255, 255)
    'global rectangle
    rWin.Left = 0
    rWin.Right = 640
    rWin.Top = 0
    rWin.Bottom = 480
'Form1.Command1.Enabled = True
End Sub
Sub UpdateLoading(perce As Integer)
With frmLoading
.Label1.Caption = perce & "%"
.Label1.Left = ((.Shape1.Left + .Shape1.Width) / 2) - (.Label1.Width / 2)
.Shape2.Width = .Shape1.Width * (perce / 100)
End With
End Sub
Sub Reinit()
    feetrect.Left = 0: feetrect.Top = 0
    GetGifInfo (Playerdetails(8))
    feetrect.Bottom = GifInfo.Height: feetrect.Right = GifInfo.Width
    Loaddirectxsurface DDFeet, Playerdetails(8)
    Loaddirectxsurface DDlegs, Playerdetails(9)
    Loaddirectxsurface DDTorso, Playerdetails(10)
    Loaddirectxsurface DDbelt, Playerdetails(11)
    Loaddirectxsurface DDarms, Playerdetails(12)
    Loaddirectxsurface DDHands, Playerdetails(13)
    Loaddirectxsurface DDCharacter, Playerdetails(8)
    Loaddirectxsurface DDHead, Playerdetails(14)
loadcharacter
End Sub
Sub LoadSettings()
music = GetSetting("JimCamel", "RPGame", "Music")
sound = GetSetting("JimCamel", "RPGame", "Sound")
plname = GetSetting("JimCamel", "RPGame", "Playername")
If sound <> "" Then Let frmOptions.Check2.Value = sound
If music <> "" Then Let frmOptions.Check1.Value = music
If plname = "" Then
Playerdetails(0) = InputBox("Please enter your name", "Enter Name", "JimCamel")
SaveSetting "JimCamel", "RPGame", "Playername", Playerdetails(0)
Else: Playerdetails(0) = plname: End If
End Sub

