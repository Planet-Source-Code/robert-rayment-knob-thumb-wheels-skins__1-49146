Attribute VB_Name = "Module1"
Option Explicit

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
Public bm As BITMAPINFO

'------------------------------------------------------------------------------
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" _
   (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
    ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

' -----------------------------------------------------------
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)
'-------------------------------------------------------

' Picture details
Public W As Long, H As Long

' Array
Public ArrMem() As Long
Public MaskMem() As Long
Public BArray() As Byte
Public comBArray() As Byte


' Booleans
Public aLoaded As Boolean
Public aArrSaved As Boolean
Public aMskSaved As Boolean

' General variables
Public STX As Long, STY As Long
Public a$
Public i As Long
Public j As Long
Public k As Long
Public resp As Long



Public Sub FillBMPStruc()
'Public ArrMem() As Long
' ArrMem(1 to W, 1 to H)
   
   With bm.bmiH
     .biSize = 40&
     .biwidth = UBound(ArrMem, 1)
     .biheight = -UBound(ArrMem, 2)
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


Public Sub GETDIB(ByVal PICIM As Long)
'Public ArrMem() As Long

' PICIM is PIC.Image - handle to PIC memory
' from which pixels will be extracted and
' stored in ArrMem().  Assumes BMPStruc already
' filled in with ArrMem parameters

Dim NewDC As Long
Dim OldH As Long

On Error GoTo DIBError

NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)

' Load color bytes to picMem
resp = GetDIBits(NewDC, PICIM, 0, UBound(ArrMem, 2), ArrMem(1, 1), bm, 1)

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC

Exit Sub
'==========
DIBError:
  MsgBox "DIB Error in GETDIBS", , "Array"
  DoEvents
  Unload Form1
  End
End Sub

Public Sub FixExtension(FSpec$, Ext$)
Dim p As Long

If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      'Ext$ = LCase$(Mid$(FSpec$, p))
      If LCase$(Mid$(FSpec$, p)) <> Ext$ Then FSpec$ = Mid$(FSpec$, 1, p) & Ext$
   End If

End Sub


