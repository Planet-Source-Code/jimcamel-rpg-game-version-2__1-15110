Attribute VB_Name = "modfilePNG"
Declare Function IsPNG Lib "filePNG.dll" (ByVal Filename As String) As Boolean
Declare Function PNGUpdateInfo Lib "filePNG.dll" (ByVal Filename As String) As Boolean
Declare Function PNGGetWidth Lib "filePNG.dll" () As Long
Declare Function PNGGetHeight Lib "filePNG.dll" () As Long
Declare Function PNGGetBPP Lib "filePNG.dll" () As Long
Declare Function PNGLoad Lib "filePNG.dll" (ByVal Filename As String, ByRef bmpDes As Byte, ByRef Palette As Any) As Boolean
Sub LoadPNG(Filename As String, ByRef pImage As ImageFile)
    Dim BytesPerLine As Long
    PNGUpdateInfo Filename
    pImage.ImageBPP = PNGGetBPP
    pImage.ImageWidth = PNGGetWidth
    pImage.ImageHeight = PNGGetHeight
    If pImage.ImageBPP = 8 Then
        BytesPerLine = TransformDIBBpl(pImage.ImageWidth, 8)
        ReDim pImage.ImageData(1 To (BytesPerLine * pImage.ImageHeight))
        ReDim pImage.ImagePalette(0 To 255)
        PNGLoad Filename, pImage.ImageData(1), pImage.ImagePalette(0)
    Else
        BytesPerLine = TransformDIBBpl(pImage.ImageWidth, 24)
        ReDim pImage.ImageData(1 To (BytesPerLine * pImage.ImageHeight))
        Erase pImage.ImagePalette
        PNGLoad Filename, pImage.ImageData(1), Nothing
    End If
End Sub
