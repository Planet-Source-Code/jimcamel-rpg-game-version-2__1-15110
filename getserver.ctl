VERSION 5.00
Begin VB.UserControl GetServer 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "getserver.ctx":0000
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "getserver.ctx":0C42
End
Attribute VB_Name = "GetServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim DownloadedText As String, Downloading As Boolean

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    Dim TextByteArray() As Byte, i As Integer, BuildText As String
    On Error Resume Next
    TextByteArray() = AsyncProp.Value
   For i = 0 To UBound(TextByteArray)
        BuildText = BuildText + Chr(TextByteArray(i))
    Next
    DownloadedText = BuildText
    Downloading = False
End Sub

Public Function DownloadText(URL As String, Optional TimeoutSeconds As Integer = 10)
   Dim OldTime As Double, CurTime As Double
    If Downloading = True Then Exit Function
    If Not Left(URL, 7) = "http://" Then
        DownloadText = "<invalid address>"
        Exit Function
    End If
    Downloading = True
    UserControl.AsyncRead URL, 2, vbNullString, vbAsyncReadForceUpdate
    OldTime = Timer
    Do
        DoEvents
        CurTime = Timer - OldTime
        If CurTime > TimeoutSeconds Then
            DownloadText = "<invalid address>"
            Exit Function
        End If
    Loop Until Downloading = False
    If DownloadedText = "" Then
        DownloadText = "<invalid address>"
    Else
        DownloadText = DownloadedText
    End If
    End Function

Private Sub UserControl_Resize()
    Size 480, 480
End Sub
