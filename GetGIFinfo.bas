Attribute VB_Name = "modGetGIFInfo"
Public Type GIFINFODATA
    Filename          As String
    Version           As String
    Width             As Long
    Height            As Long
    Numbercol         As Long
    Palette(255)      As Long
    Transparent       As Boolean
    Transindex        As Long
  End Type
  Global filetype
  Global GifInfo As GIFINFODATA

Public Declare Function IsPNG Lib "filePNG.dll" (ByVal Filename As String) As Boolean
Public Declare Function PNGUpdateInfo Lib "filePNG.dll" (ByVal Filename As String) As Boolean
Public Declare Function PNGGetWidth Lib "filePNG.dll" () As Long
Public Declare Function PNGGetHeight Lib "filePNG.dll" () As Long
Public Declare Function PNGGetBPP Lib "filePNG.dll" () As Long
Public Declare Function PNGLoad Lib "filePNG.dll" (ByVal Filename As String, ByRef bmpDes As Byte, ByRef Palette As Any) As Boolean
'Sub LoadPNG(Filename As String, ByRef pImage As ImageFile)
'    Dim BytesPerLine As Long
'    PNGUpdateInfo Filename
'    MsgBox PNGGetBPP
 '   MsgBox PNGGetWidth
 '   MsgBox PNGGetHeight
 ''   pImage.ImageBPP = PNGGetBPP
 '   pImage.ImageWidth = PNGGetWidth
 '   pImage.ImageHeight = PNGGetHeight
 '   If pImage.ImageBPP = 8 Then
 '       BytesPerLine = TransformDIBBpl(pImage.ImageWidth, 8)
  '      ReDim pImage.ImageData(1 To (BytesPerLine * pImage.ImageHeight))
  '      ReDim pImage.ImagePalette(0 To 255)
  ''      PNGLoad Filename, pImage.ImageData(1), pImage.ImagePalette(0)
  '  Else
  '      BytesPerLine = TransformDIBBpl(pImage.ImageWidth, 24)
  '      ReDim pImage.ImageData(1 To (BytesPerLine * pImage.ImageHeight))
  '      Erase pImage.ImagePalette
  '      PNGLoad Filename, pImage.ImageData(1), Nothing
  '  End If
'End Sub

  Public Sub Checkfiletype(ByVal Filename As String)
    Dim bChar As Byte
    Dim Header As String
    fnum = FreeFile
    Open Filename For Binary Access Read As #fnum
    For i = 0 To 5
        Get #fnum, , bChar
        Header = Header + Chr(bChar)
    Next i
If InStr(1, Header, "PNG") > 0 Then
filetype = "PNG"
ElseIf InStr(1, Header, "GIF") > 0 Then
filetype = "GIF"
Else
filetype = ""
End If
          Close #fnum
          End Sub
Public Sub GetGifInfo(ByVal Filename As String)
'declare temporary variables
    Dim bChar As Byte
    Dim i As Integer
    Dim DotPos As Integer
    Dim Header As String
    Dim blExit As Boolean
    Dim a As String, b As String, c As String, d As String
    Dim ImgWidth As Integer
    Dim ImgHeight As Integer
    Dim Transcol As Integer
    Dim ImgSize As String
    Dim fnum As Integer
    Dim ver As String
    Dim Numbercol As Long
    Dim pal(255) As Long
    Dim trans As Boolean
    Dim transin As Long
    On Error Resume Next
    
    fnum = FreeFile
    'Open the file as an availible filenumber
    Open Filename For Binary Access Read As #fnum

    ImgSize = LOF(fnum) / 1024

    DotPos = InStr(ImgSize, ",")
    ImgSize = Left(ImgSize, DotPos - 1)

'read the first 5 bytes of data, which is the description
    For i = 0 To 5
        Get #fnum, , bChar
        Header = Header + Chr(bChar)
    Next i

Let ver = Right(Header, 3)

    'the next 2 bytes indicate the image width
    Get #fnum, , bChar
    a = a + Chr(bChar)
    Get #fnum, , bChar
    a = a + Chr(bChar)
    ImgWidth = CInt(Asc(Left(a, 1)) + 256 * Asc(Right(a, 1)))
    
    'the next 2 bytes indicate image height
    Get #fnum, , bChar
    b = b + Chr(bChar)
    Get #fnum, , bChar
    b = b + Chr(bChar)
    ImgHeight = CInt(Asc(Left(b, 1)) + 256 * Asc(Right(b, 1)))
    
    'the next 3 bytes indicate the number of colours
    Get #fnum, , bChar
    c = c + Chr(bChar)
    Get #fnum, , bChar
    c = c + Chr(bChar)
    Get #fnum, , bChar
    c = c + Chr(bChar)
    
    'Here is a problem. I'm not sure how exactly to read this data
    'when asc(c) = 179 there's 16 colours, when it's 247, it's 256 colours
    'If you know how to work this out, please email me
    If Asc(c) = 179 Then c = 16
    If Asc(c) = 247 Then c = 256
    Numbercol = c
    
    Dim red As Long
    Dim green As Long
    Dim blue As Long
    
    'The next block tells us the palette colours
    'Each colour has 3 bytes, a red, green and blue colour
    'So the total no of bytes is = Number of colours x 3
    
    For i = 0 To Numbercol - 1
    Get #fnum, , bChar
    red = bChar
    Get #fnum, , bChar
    green = bChar
    Get #fnum, , bChar
    blue = bChar
    'Put all the information into the pal placeholder as long values
    pal(i) = RGB(red, green, blue)
        Next i
    'The 4th byte after the palette is a boolean flag on whether there
    'is transparency or not. I'm not sure if this is present in 87a
    'GIF files
    
    Get #fnum, 17 + (3 * Numbercol), bChar
    c = bChar
    If c <> 0 And c <> 1 Then
    transtest = 1
    ElseIf c = 1 Then Let trans = True
    Else: Let trans = False:
    End If
    
    'The 3rd byte after the transparency flag is the palette number which
    'is transparent
    
    Get #fnum, 20 + (3 * Numbercol), bChar
    c = bChar
    transin = c
    Close #fnum
If transtest = 1 And transin <> 0 Then Let trans = True
    'Copy all the information into the GifInfo type for later reference
    With GifInfo
         .Filename = Filename
         .Version = ver
         .Width = ImgWidth
         .Height = ImgHeight
         .Numbercol = Numbercol
    For i = 0 To Numbercol - 1
         .Palette(i) = pal(i)
    Next i
         .Transparent = trans
  
        If trans = True Then .Transindex = transin
    End With
'MsgBox Filename & ":" & pal(transin)
End Sub

