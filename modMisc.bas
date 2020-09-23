Attribute VB_Name = "modMisc"
'***********************************************************
'* Name        : Custom Image Format Sample                *
'* Description : Shows how to easily load and save custom  *
'*               Image formats using DIB Functions. This   *
'*               Allows you to quickly Get/Put your image  *
'*               From/Onto any Device Context.             *
'* Date        : 11-May-2000                               *
'* Author      : Adam Hoult                                *
'* e-Mail      : admin@daedalusd.com                       *
'* Website     : http://www.daedalusd.com/vbgaming         *
'***********************************************************

Option Explicit

'Declarations
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Sub SetMemory Lib "kernel32" Alias "RtlFillMemory" (pDst As Any, Fillchar As Any, ByVal ByteLen As Long)

'Constants
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

'Types
'We share this type with the API, we will use it to store our
'palettes also.
Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
'This is our Custom Image File Format.
Public Type ImageFile
        ImageWidth As Long          'Width of Image
        ImageHeight As Long         'Height of Image
        ImageBPP As Byte            'Bits Per Pixel (24bpp or 8bpp)
        ImagePalette() As RGBQUAD   'Store our palette (for 8bpp only)
        ImageData() As Byte         'Series of Bytes (R-G-B for 24bpp or Palette Index for 8bpp)
        Handle As OLE_HANDLE
End Type

'The below types are used by the win32 API
Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(0 To 255) As RGBQUAD
End Type

Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Variables
Public Img As ImageFile        'This is our Image object used to load/save etc.

Dim bmi As BITMAPINFO                'Stores the DIB information
Function TransformDIBBpl(Value As Long, BPP As Long) As Long
    'This function is used to transform an actual width
    'to the number of bytes per line needed for windows
    'to use as a DIB.
    Select Case BPP
        Case 24:
            TransformDIBBpl = (Value * 3 + 3) And &HFFFFFFFC
        Case 8:
            TransformDIBBpl = (Value + 3) And &HFFFC
    End Select
End Function
Sub LoadImage(Filename As String, ByRef pImage As ImageFile, ByRef pBox As PictureBox)
    If IsBMP(Filename) = True Then
        LoadBMP Filename, pImage
    ElseIf IsJPEG(Filename) Then
        LoadJPEG Filename, pImage, pBox
    ElseIf IsGIF(Filename) Then
        LoadGIF Filename, pImage, pBox
    ElseIf IsPNG(Filename) Then
        LoadPNG Filename, pImage
    ElseIf IsPSD(Filename) Then
        LoadPSD Filename, pImage
    ElseIf IsTGA(Filename) Then
        LoadTGA Filename, pImage
    ElseIf IsTIFF(Filename) Then
        LoadTIFF Filename, pImage
    ElseIf IsPCX(Filename) Then
        LoadPCX Filename, pImage
    ElseIf IsLBM(Filename) Then
        LoadLBM Filename, pImage
    ElseIf IsVGA(Filename) Then
        LoadVGA Filename, pImage
    End If
End Sub
Function ReturnExtension(ByVal pName As String) As String
'Function Name : ReturnExtension
'Called By     : General
'Description   : Returns the last real extension in a filename
'Returns       : String Extension
'Parameters    : pName, Filename to extract extension from.

    pName = Trim(pName)
    ' Keep getting rid of everything before the dots (and the dot)
    ' i.e If your filename was c:\newfolder.001\folder.002\hello.jpg
    ' The Stages would be
    'c:\newfolder.001\folder.002\hello.jpg
    '001\folder.002\hello.jpg
    '002\hello.jpg
    'jpg
    If InStr(pName, ".") = 0 Then ReturnExtension = "": Exit Function
    Do Until InStr(pName, ".") = 0
        pName = Mid(pName, InStr(pName, ".") + 1)
    Loop
    'Return the Extension in uppercase
    ReturnExtension = UCase(pName)
    
End Function

Sub DrawImage(DC As Long, ByRef pImage As ImageFile)
    On Error Resume Next
    bmi.bmiHeader.biSize = 40                       'Simply the size of the header type
    bmi.bmiHeader.biWidth = pImage.ImageWidth        'Width of Image
    bmi.bmiHeader.biHeight = -pImage.ImageHeight    'Height of image (reversed)
    bmi.bmiHeader.biPlanes = 1                      'The image has 1 plane
    bmi.bmiHeader.biBitCount = pImage.ImageBPP       'How many Bits Per Pixel.
    bmi.bmiHeader.biCompression = 0                 'No Compression used.
    If pImage.ImageBPP = 8 Then CopyMemory bmi.bmiColors(0), pImage.ImagePalette(0), 256 * 4
    'Now take this information and copy the data onto the main pictures HDC.
    'Remember that StretchDIBits also draws from the bottom left corner upwards, so if our
    'data is stored from top to bottom, then we must negate the height value which will
    'flip the image.
    StretchDIBits DC, 0, 0, pImage.ImageWidth, pImage.ImageHeight, 0, 0, pImage.ImageWidth, pImage.ImageHeight, pImage.ImageData(1), bmi, DIB_RGB_COLORS, vbSrcCopy
End Sub
Sub LoadCustomImage(Filename As String)
'Description : Loads our custom image from the file specified, and copies the Bits
'              onto our main picture.
'Date        : 11-May-2000
    Open Filename For Binary Access Read As #1
        Get #1, , Img                           'Read in our UDT image
    Close #1
End Sub
Sub SaveCustomImage(Filename As String)
'Description : Reads the Bits from our main picture and copies them into
'              our custom image. Then saves to the file specified.
'Date        : 11-May-2000
    If Dir(Filename) <> "" Then Kill Filename
    Open Filename For Binary Access Write As #1
        Put #1, , Img
    Close #1
End Sub

Public Function LShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long
    LShift = lValue * (2 ^ lNumberOfBitsToShift)
End Function

Public Function RShift(ByVal lValue As Long, ByVal lNumberOfBitsToShift As Long) As Long
    RShift = lValue \ (2 ^ lNumberOfBitsToShift)
End Function
Function fgetWord(filenumber As Integer, intel As Boolean) As Long
    Dim i As Long, j As Long, b0 As Byte, b1 As Byte
    Get #filenumber, , b0
    Get #filenumber, , b1
    i = b0 And 255
    j = b1 And 255
    If intel = True Then fgetWord = (i + (LShift(j, 8))) Else fgetWord = (LShift(i, 8) + j)
End Function
Function fgetLong(filenumber As Integer, intel As Boolean) As Long
    Dim i As Long, j As Long, k As Long, l As Long
    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
    Get #filenumber, , b0
    Get #filenumber, , b1
    Get #filenumber, , b2
    Get #filenumber, , b3
    i = b0 And 255
    j = b1 And 255
    k = b2 And 255
    l = b3 And 255
    If intel = True Then
        fgetLong = (i + LShift(j, 8) + LShift(k, 16) + LShift(l, 24))
    Else
        fgetLong = (l + LShift(k, 8) + LShift(j, 16) + LShift(i, 24))
    End If
End Function


