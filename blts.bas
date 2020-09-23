Attribute VB_Name = "Bltbas"
Sub bltmapclipped()
mapplx = 10
mapply = 10
Dim r1 As RECT
With r1
.Left = rWin.Left
.Top = rWin.Top
.Right = rWin.Right
.Bottom = rWin.Bottom
End With
If plx > 10 Then
r1.Left = (plx - 10) * 16
r1.Right = r1.Left + (rWin.Right - rWin.Left)
mapplx = 10
Else
mapplx = plx
End If
If (ddmapsize.Right / 48) - plx <= 10 Then
r1.Right = ddmapsize.Right / 2
r1.Left = r1.Right - (rWin.Right - rWin.Left)
mapplx = Round(plx - ((ddmapsize.Right / 48) - ((rWin.Right - rWin.Left) / 48)) + 6.9, 1)
End If
If ply >= 7.5 Then
r1.Top = (ply - 7.5) * 16
r1.Bottom = r1.Top + (rWin.Bottom - rWin.Top)
mapply = 7.5
Else
mapply = ply
End If
If (ddmapsize.Bottom / 48) - ply <= 7.5 Then
r1.Bottom = ddmapsize.Bottom / 2.18
r1.Top = r1.Bottom - (rWin.Bottom - rWin.Top)
mapply = ply - ((ddmapsize.Bottom / 48) - ((rWin.Bottom - rWin.Top) / 48)) + 5
End If
'MsgBox ddmapsize.Right
'Debug.Print r1.Left & ":" & r1.Right
DDmapbuffer.Blt rWin, DDMap, r1, DDBLT_WAIT
DDmapbuffer.Blt rWin, DDOppbuffer, r1, DDBLT_KEYSRC
'If connected = True And plonline > 1 Then
End Sub
Sub Blt()
If GameForm.Visible = False Then Exit Sub
If frmLogin.Visible = True Then Exit Sub
On Error Resume Next
Dim rw As RECT: Dim rhead As RECT: Dim rheadpos As RECT: Dim cw As Integer: Dim ch As Integer
Dim postemp(1) As Integer: Dim scrtemp(1) As Integer
Call DX.GetWindowRect(GameForm.DXPic.hwnd, rw)
If edit <> True Then
bltmapclipped
'MsgBox connected & ":" & plonline
If connected = True And plonline > 1 Then bltopp
Call bltplayer1(Imgy, Imgx)
r1.Top = 50: r1.Bottom = 530: r1.Left = 0: r1.Right = 640
DDGetReady.Blt r1, DDmapbuffer, rWin, DDBLT_WAIT
BltGUI
If quitmenu = True Then bltmenu
End If
bltNPC
DDPrimSurf.Blt rw, DDGetReady, rWin, DDBLT_WAIT
End Sub
Sub bltNPC()
'For i = 1 To NPCno

'Next i
End Sub
Sub bltmouseout(mouseout As Integer)
Dim rw As RECT
Dim r5 As RECT: Dim r6 As RECT:
Call DX.GetWindowRect(GameForm.DXPic.hwnd, rw)
r5.Left = GUI(mouseout, 2): r5.Top = GUI(mouseout, 3): r5.Right = GUI(mouseout, 4): r5.Bottom = GUI(mouseout, 5)
r6.Top = GUI(mouseout, 7): r6.Bottom = GUI(mouseout, 9): r6.Left = GUI(mouseout, 6): r6.Right = GUI(mouseout, 8)
rw.Top = rw.Top + r6.Top: rw.Bottom = rw.Top + (r6.Bottom - r6.Top): rw.Left = rw.Left + r6.Left: rw.Right = rw.Left + (r6.Right - r6.Left)
DDGetReady.Blt r6, DDGUI(GUI(mouseout, 12)), r5, DDBLT_KEYSRC
DDPrimSurf.Blt rw, DDGetReady, r6, DDBLT_WAIT
End Sub

Sub bltmouseover(mouseover As Integer)
Dim rw As RECT
Dim r5 As RECT: Dim r6 As RECT:
Call DX.GetWindowRect(GameForm.DXPic.hwnd, rw)
For i = 1 To guinumber
If GUI(i, 0) = GUI(mouseover, 10) Then
r5.Left = GUI(i, 2): r5.Top = GUI(i, 3): r5.Right = GUI(i, 4): r5.Bottom = GUI(i, 5)
r6.Top = GUI(i, 7): r6.Bottom = GUI(i, 9): r6.Left = GUI(i, 6): r6.Right = GUI(i, 8)
rw.Top = rw.Top + r6.Top: rw.Bottom = rw.Top + (r6.Bottom - r6.Top): rw.Left = rw.Left + r6.Left: rw.Right = rw.Left + (r6.Right - r6.Left)
DDGetReady.Blt r6, DDGUI(GUI(i, 12)), r5, DDBLT_KEYSRC
DDPrimSurf.Blt rw, DDGetReady, r6, DDBLT_WAIT
End If
Next i
mouseov = mouseover
End Sub
Sub bltmenu()
Dim r5 As RECT: Dim r6 As RECT: Dim rw As RECT
Call DX.GetWindowRect(GameForm.DXPic.hwnd, rw)
If quitmenu = True Then
For i = 1 To guinumber
If GUI(i, 0) = "MenuButton" Then GoTo 1
If Left(GUI(i, 0), 4) = "Menu" Then
r5.Left = GUI(i, 2): r5.Top = GUI(i, 3): r5.Right = GUI(i, 4): r5.Bottom = GUI(i, 5)
r6.Left = GUI(i, 6): r6.Top = GUI(i, 7): r6.Right = GUI(i, 8): r6.Bottom = GUI(i, 9)
DDGetReady.Blt r6, DDGUI(GUI(i, 12)), r5, DDBLT_KEYSRC
End If
1 Next i
End If
End Sub
Sub BltGUI()
Dim r5 As RECT: Dim r6 As RECT: Dim r7 As RECT: Dim R8 As RECT:
For i = 1 To guinumber
If GUI(i, 11) = "True" Then
r5.Left = GUI(i, 2): r5.Top = GUI(i, 3): r5.Right = GUI(i, 4): r5.Bottom = GUI(i, 5)
r6.Left = GUI(i, 6): r6.Top = GUI(i, 7): r6.Right = GUI(i, 8): r6.Bottom = GUI(i, 9)
DDGetReady.Blt r6, DDGUI(GUI(i, 12)), r5, DDBLT_KEYSRC
End If
1 Next i

For i = 1 To guinumber
If GUI(i, 0) = "Health" Or GUI(i, 0) = "health" Then Let a = i
If GUI(i, 0) = "NoHealth" Then Let b = i
Next i

For i = 0 To Playerdetails(2) - 1
r5.Left = GUI(a, 2): r5.Top = GUI(a, 3): r5.Right = GUI(a, 4): r5.Bottom = GUI(a, 5)
r6.Left = (GUI(a, 6) + (GUI(a, 15) * i)): r6.Top = GUI(a, 7): r6.Right = r6.Left + (GUI(a, 4) - GUI(a, 2)): r6.Bottom = GUI(a, 9)
DDGetReady.Blt r6, DDGUI(GUI(a, 12)), r5, DDBLT_KEYSRC
Next i
If Playerdetails(2) < 10 Then
For i = Playerdetails(2) To 9
r5.Left = GUI(b, 2): r5.Top = GUI(b, 3): r5.Right = GUI(b, 4): r5.Bottom = GUI(b, 5)
r6.Left = (GUI(a, 6) + (GUI(a, 15) * i)): r6.Top = GUI(a, 7): r6.Right = r6.Left + (GUI(a, 4) - GUI(a, 2)): r6.Bottom = GUI(a, 9)
DDGetReady.Blt r6, DDGUI(GUI(b, 12)), r5, DDBLT_KEYSRC
Next i
End If

Dim c(4) As Integer
For i = 1 To guinumber
If GUI(i, 0) = "box1" Then Let c(1) = i
If GUI(i, 0) = "box2" Then Let c(2) = i
If GUI(i, 0) = "box3" Then Let c(3) = i
If GUI(i, 0) = "box4" Then Let c(4) = i
If GUI(i, 0) = "Numbers" Then Let e = i
Next i
Let f = CInt((GUI(e, 4) - GUI(e, 2)) / 10)

For j = 1 To 4
For i = Len(Playerdetails(j + 2)) To 1 Step -1
r6.Left = GUI(e, 2) + (Mid(Playerdetails(j + 2), i, 1) * f): r6.Top = GUI(e, 3): r6.Bottom = GUI(e, 5): r6.Right = r6.Left + f
r7.Left = GUI(c(j), 8) - ((Len(Playerdetails(j + 2)) - (i - 1.5)) * f): r7.Top = GUI(c(j), 7) + 1.5: r7.Bottom = r7.Top + GUI(e, 5) - GUI(e, 3): r7.Right = r7.Left + f
DDGetReady.Blt r7, DDGUI(GUI(e, 12)), r6, DDBLT_KEYSRC
Next i
Next j

For i = 1 To guinumber
If GUI(i, 0) = "StaminaFill" Then
r5.Left = GUI(i, 2): r5.Top = GUI(i, 3): r5.Right = GUI(i, 4): r5.Bottom = GUI(i, 5)
r6.Left = GUI(i, 6): r6.Top = GUI(i, 7): r6.Right = GUI(i, 6) + CInt((GUI(i, 8) - GUI(i, 6)) * (Playerdetails(7) / 100)): r6.Bottom = GUI(i, 9)
DDGetReady.Blt r6, DDGUI(GUI(i, 12)), r5, DDBLT_KEYSRC
End If
Next i

End Sub

Sub bltmap()
Dim rtemp As RECT
Dim Maprect As RECT
For mainheight = 0 To ddmapsize.Bottom / 16
For mainwidth = 0 To ddmapsize.Right / 16
rtemp.Left = mainwidth * 16
rtemp.Right = (mainwidth + 1) * 16
rtemp.Top = mainheight * 16
rtemp.Bottom = (mainheight + 1) * 16
Maprect.Left = map(mainwidth, mainheight, 1) * 16: Maprect.Top = map(mainwidth, mainheight, 2) * 16: Maprect.Right = Maprect.Left + 16: Maprect.Bottom = Maprect.Top + 16
DDMap.Blt rtemp, DDBackround, Maprect, DDBLT_WAIT
Next mainwidth
Next mainheight
'MsgBox rtemp.Right & ":" & rtemp.Bottom
End Sub
Sub bltopp()
Dim rw As RECT:
Dim rhead As RECT: Dim rheadpos As RECT: Dim cw As Integer: Dim ch As Integer
'MsgBox "plonline:" & plonline
For j = 1 To plonline
action = oppdetail(j, 4)
For i = 1 To 2
r1.Left = characterdata(action, i, oppdetail(i, 3), 3):
r1.Right = characterdata(action, i, oppdetail(i, 3), 3) + characterdata(action, i, oppdetail(i, 3), 5)
r1.Top = characterdata(action, i, oppdetail(i, 3), 4):
r1.Bottom = characterdata(action, i, oppdetail(i, 3), 4) + characterdata(action, i, oppdetail(i, 3), 6)
r2.Left = (oppdetail(j, 1) * 32) + characterdata(action, i, oppdetail(i, 3), 1):
r2.Right = (oppdetail(j, 1) * 32) + characterdata(action, i, oppdetail(i, 3), 1) + characterdata(action, i, oppdetail(i, 3), 5)
r2.Top = (oppdetail(j, 2) * 32) + characterdata(action, i, oppdetail(i, 3), 2):
r2.Bottom = (oppdetail(j, 2) * 32) + characterdata(action, i, oppdetail(i, 3), 2) + characterdata(action, i, oppdetail(i, 3), 6)
'MsgBox i & ":" & oppdetail(j, 0) & vbCr & r1.Left & ":" & r1.Top & ":" & r1.Right & ":" & r1.Bottom & vbCr & r2.Left & ":" & r2.Top & ":" & r2.Right & ":" & r2.Bottom
If i = 1 Then DDOppbuffer.Blt r2, DDOpponents(j), r1, DDBLT_KEYSRC
If i = 2 Then DDOppbuffer.Blt r2, DDOpphead(j), r1, DDBLT_KEYSRC
Next i
DDOppbuffer.DrawText (oppdetail(j, 1) * 32 - ((Len(oppdetail(j, 0)) / 2) * 6)), (oppdetail(j, 2) + 0.8) * 32, oppdetail(j, 0), False
If oppdetail(j, 15) <> "" Then DDOppbuffer.DrawText (oppdetail(j, 1) * 32 - ((Len(oppdetail(j, 6)) / 2) * 6)), (oppdetail(j, 2) - 1) * 32, oppdetail(j, 15), False
Next j
End Sub
Sub bltplayer1(action As Integer, Direction As Integer)
Dim rw As RECT:
Dim rhead As RECT: Dim rheadpos As RECT: Dim cw As Integer: Dim ch As Integer
For i = 1 To 2
r1.Left = characterdata(action, i, Direction, 3):
r1.Right = characterdata(action, i, Direction, 3) + characterdata(action, i, Direction, 5)
r1.Top = characterdata(action, i, Direction, 4):
r1.Bottom = characterdata(action, i, Direction, 4) + characterdata(action, i, Direction, 6)
r2.Left = (mapplx * 32) + characterdata(action, i, Direction, 1):
r2.Right = (mapplx * 32) + characterdata(action, i, Direction, 1) + characterdata(action, i, Direction, 5)
r2.Top = (mapply * 32) + characterdata(action, i, Direction, 2):
r2.Bottom = (mapply * 32) + characterdata(action, i, Direction, 2) + characterdata(action, i, Direction, 6)
If i = 1 Then DDmapbuffer.Blt r2, DDCharacter, r1, DDBLT_KEYSRC
If i = 2 Then DDmapbuffer.Blt r2, DDHead, r1, DDBLT_KEYSRC
Next i
DDmapbuffer.DrawText (mapplx * 32 - ((Len(Playerdetails(0)) / 2) * 5)), (mapply * 32) + characterdata(action, 1, Direction, 6) - 5, Playerdetails(0), False
DDmapbuffer.DrawText (mapplx * 32 - ((Len(Playerdetails(1)) / 2) * 5)), (mapply * 32) - characterdata(action, 2, Direction, 6) + 5, Playerdetails(1), False
End Sub

