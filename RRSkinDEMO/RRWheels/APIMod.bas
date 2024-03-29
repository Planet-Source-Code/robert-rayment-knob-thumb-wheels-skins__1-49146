Attribute VB_Name = "APIMod"
' APIMod.bas
Option Explicit


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

'Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'
'Public Declare Function SetPixelV Lib "gdi32" _
'(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'
'Public Declare Function SetCursorPos Lib "user32" _
'(ByVal X As Long, ByVal Y As Long) As Long
'
'Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
' ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
'Public Const FLOODFILLSURFACE = 1

' -----------------------------------------------------------

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
'   Colors(0 To 255) As RGBQUAD
End Type
Public bmStart As BITMAPINFO
Public bmStrip As BITMAPINFO

' For transferring drawing in memory to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Const DIB_RGB_COLORS = 1 '  uses System

'-------------------------------------------------------------------------

'Declare Function GetStretchBltMode Lib "gdi32" _
'(ByVal hdc As Long) As Long

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4

' Use:-
'  oldMode = GetStretchBltMode( pic.hdc)
'
'  If SetStretchBltMode(pic.hdc, HALFTONE) = 0 Then
'     MsgBox "SetStretchBltMode error ", vbCritical, " "
'     End
'  End If
'
'  SetStretchBltMode pic.hdc, OldMode
'-------------------------------------------------------------------------

'------------------------------------------------------------------------------

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'------------------------------------------------------------------------------
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" _
   (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
    ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


Public Sub FillBMPStruc(bMem() As Byte, bmm As BITMAPINFO)
'Public Mem() As Byte
' Mem(4,1 to W, 1 to H)
   
   With bmm.bmiH
     .biSize = 40&
     .biwidth = UBound(bMem, 2)
     .biheight = -UBound(bMem, 3)
     .biPlanes = 1
     .biBitCount = 32
     .biCompression = 0&
     .biSizeImage = Abs(.biwidth) * Abs(.biheight) * 4
     .biXPelsPerMeter = 0&
     .biYPelsPerMeter = 0&
     .biClrUsed = 0&
     .biClrImportant = 0&
   End With
End Sub


Public Sub GETDIB(ByVal PICIM As Long, bMem() As Byte, bmm As BITMAPINFO)
' PICIM is picbox.Image - handle to picbox memory
' from which pixels will be extracted and
' stored in Mem().  Assumes BMPStruc already
' filled in with picMem parameters

Dim NewDC As Long
Dim OldH As Long

On Error GoTo DIBError

   NewDC = CreateCompatibleDC(0&)
   OldH = SelectObject(NewDC, PICIM)
   
   ' Load color bytes to picMem
   GetDIBits NewDC, PICIM, 0, UBound(bMem, 3), bMem(1, 1, 1), bmm, 1
   
   ' Clear mem
   SelectObject NewDC, OldH
   DeleteDC NewDC

Exit Sub
'==========
DIBError:
  MsgBox "DIB Error in GETDIBS", vbCritical, " "
  DoEvents
  Unload Form1
  End
End Sub

