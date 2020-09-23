Attribute VB_Name = "Module1"
' Module1.bas

Option Explicit


'Public Type POINTAPI
'        x As Long
'        y As Long
'End Type
'Public PT As POINTAPI
'
'Public Declare Function SetBrushOrgEx Lib "gdi32" _
'(ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
'-------------------------------------------------------------------------

'' Structures for StretchDIBits
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
'------------------------------------------------------------------------------

' For transferring drawing in memory to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Long, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Const DIB_RGB_COLORS = 1 '  uses System
'------------------------------------------------------------------------------

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'------------------------------------------------------------------------------
' Resizing APIs
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTBOTTOMRIGHT = 17


'Public Const HTBOTTOM = 15
'Public Const HTBORDER = 18
'Public Const HTBOTTOMLEFT = 16
'Public Const HTBOTTOMRIGHT = 17
'Public Const HTCAPTION = 2
'Public Const HTCLIENT = 1
'Public Const HTERROR = (-2)
'Public Const HTGROWBOX = 4
'Public Const HTHSCROLL = 6
'Public Const HTLEFT = 10
'Public Const HTMAXBUTTON = 9
'Public Const HTMENU = 5
'Public Const HTMINBUTTON = 8
'Public Const HTNOWHERE = 0
'Public Const HTREDUCE = HTMINBUTTON
'Public Const HTRIGHT = 11
'Public Const HTSIZE = HTGROWBOX
'Public Const HTSIZEFIRST = HTLEFT
'Public Const HTSIZELAST = HTBOTTOMRIGHT
'Public Const HTSYSMENU = 3
'Public Const HTTOP = 12
'Public Const HTTOPLEFT = 13
'Public Const HTTOPRIGHT = 14
'Public Const HTTRANSPARENT = (-1)
'Public Const HTVSCROLL = 7
'Public Const HTZOOM = HTMAXBUTTON


Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'------------------------------------------------------------------------------
' Ahmad Marwan Mami's notation (PSC CodeId=42081)
Public Declare Function SWR Lib "user32" Alias "SetWindowRgn" _
   (ByVal hWnd As Long, ByVal hrgn As Long, _
    ByVal bRedraw As Boolean) As Long
    
Public Declare Function CRR Lib "gdi32" Alias "CreateRoundRectRgn" _
  (ByVal XTL As Long, ByVal YTL As Long, _
   ByVal XBR As Long, ByVal YBR As Long, _
   ByVal EW As Long, ByVal EH As Long) As Long
   
   ' XTL,YTL  XBR,YBR  Top left & Bottom right coords of rectangle
   ' EW,EH  width & height of ellipse used to create corners
'-------------------------------------------------------------------------

'Declare Function GetStretchBltMode Lib "gdi32" _
'(ByVal hdc As Long) As Long

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

Public Const COLORONCOLOR = 3
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

'Public Declare Function GetDeviceCaps Lib "gdi32" _
'() '(ByVal hdc As Long, ByVal nIndex As Long) As Long
'
'Public Const VERTRES = 10
'-------------------------------------------------------------------------

Public Declare Function GetSystemMetrics Lib "user32" _
(ByVal nIndex As Long) As Long

Public Const SM_CXSCREEN = 0  ' Screen Width
Public Const SM_CYSCREEN = 1  ' Screen Height
'Public Const SM_CYCAPTION = 4 ' Height of window caption
'Public Const SM_CYMENU = 15   ' Height of menu
'Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
'Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)
'-------------------------------------------------------------------------

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)
'-------------------------------------------------------------------------

' Windows API - For playing WAV files   NB Win98
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
        
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4    ' lpszSoundName points to data in memory
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_PURGE = &H40

'- SND_SYNC specifies that the sound is played synchronously and the
'  function does not return until the sound ends.

'- SND_ASYNC specifies that the sound is played asynchronously and the
'  function returns immediately after beginning the sound.

'- SND_NODEFAULT specifies that if the sound cannot be found, the
'  function returns silently without playing the default sound.

'- SND_MEMORY sound played from memory (eg a String)

'- SND_LOOP specifies that the sound will continue to play continuously
'  until PlaySound is called again with the lpszSoundName$ parameter
'  set to null. You must also specify the SND_ASYNC flag to loop sounds.

'- SND_NOSTOP specifies that if a sound is currently playing, the
'  function will immediately return False without playing the requested
'  sound.

'_ SND_PURGE Stop playback


'     To hold data from Resource WAV file
'EG   Public WAVData As String
'EG   ' Get wav data
'     WAVData = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)

'EG   sndPlaySound WAVData, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY

'EG   sndPlaySound "", SND_PURGE

Public WAVSWITCH As String
Public WAVSLIDE As String
Public WAVTIC As String
Public WAVTIC2 As String

'-------------------------------------------------------------------------

Public bm As BITMAPINFO
Public bmslider As BITMAPINFO
Public bmscale As BITMAPINFO
Public bmwatchit As BITMAPINFO
Public bmpvscale As BITMAPINFO
Public bmpvslider As BITMAPINFO
Public bmwheel As BITMAPINFO
Public bmcross As BITMAPINFO
Public bmswitch As BITMAPINFO

Public ORGFW As Long ' Org form width and height
Public ORGFH As Long
Public XMag As Single
Public YMag As Single
Public xmagslide As Single
Public ymagslide As Single

' Picture size
Public W As Long, H As Long

Public w1 As Long, h1 As Long    ' Slider & Slider Mask size
Public w2 As Long, h2 As Long    ' Scale size

Public w4 As Long, h4 As Long    ' WatchIt & WatchIt Mask size
Public w5 As Long, h5 As Long    ' pvScale & pvScale Mask size
Public w6 As Long, h6 As Long    ' pvSlider & pvSlider Mask size

Public w7 As Long, h7 As Long    ' picWheel size

Public w8 As Long, h8 As Long    ' picCross size

Public w9 As Long, h9 As Long    ' picSwitch size

' Arrays
Public BArray() As Byte
Public uncomBArray() As Byte
Public ByteLen As Long
Public ByteLen2 As Long

Public ArrMem() As Long

Public SliderMem() As Long
Public SliderMaskMem() As Long
Public ScaleMem() As Long
Public WatchItMem() As Long
Public WatchItMaskMem() As Long

Public pvScaleMem() As Long
Public pvScaleMaskMem() As Long
Public pvSliderMem() As Long
Public pvSliderMaskMem() As Long

Public WheelMem() As Long

Public CrossUpMem() As Long
Public CrossDnMem() As Long

Public SwitchUpMem() As Long
Public SwitchUpMaskMem() As Long
Public SwitchDnMem() As Long
Public SwitchDnMaskMem() As Long

' Wheel
Public xcw As Single    ' Wheel center
Public ycw As Single    ' Wheel center
Public zprevAngle As Single
Public zAngle As Single
Public Const pi# = 3.1415927

' Files
Public PathSpec$, CurrentPath$
Public FileSpec$

' Booleans
Public aLoaded As Boolean
Public aResize As Boolean
Public aChecked As Boolean
Public aTile As Boolean

' Switch
Public SwitchYPrev As Long
Public aSwitch2 As Boolean

' General variables
Public STX As Long, STY As Long
'Public a$
Public i As Long
Public j As Long
'Public k As Long
'Public resp As Long

Public Sub FillStruc(bmm As BITMAPINFO, Arr() As Long, Num As Long)
   With bmm.bmiH
     .biSize = 40&
     .biwidth = UBound(Arr, 1)
     .biheight = -UBound(Arr, 2)
     .biPlanes = Int(1)
     .biBitCount = Int(32)
     .biCompression = 0&
     .biSizeImage = Abs(.biwidth * .biheight) * 4
     .biXPelsPerMeter = 0&
     .biYPelsPerMeter = 0&
     .biClrUsed = 0&
     .biClrImportant = 0&
   End With
End Sub

Public Sub LOAD_ARRAYS(aFileErrors As Boolean)
Dim N As Long  ' For error checking

   On Error GoTo FileErr
   aFileErrors = False
   
   '===================================================
   
   
   ' 1 Special .arz for picSlider
   '------------------------------
   N = 1
   BArray = LoadResData("SLIDERBA", "CUSTOM")
   INFLATE w1, h1, 1
   ReDim SliderMem(1 To w1, 1 To h1)
   CopyMemory SliderMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   FillStruc bmslider, SliderMem(), 1
   '------------------------------
   ' 1 Special .msz for picSlider
   '------------------------------
   N = 11
   BArray = LoadResData("SLIDERBM", "CUSTOM")
   DoEvents
   Call INFLATE(w1, h1, 11)
   ReDim SliderMaskMem(1 To w1, 1 To h1)
   CopyMemory SliderMaskMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   
   '===================================================
   ' 2 Special .arz for picScale
   '------------------------------
   N = 2
   BArray = LoadResData("SCALEDARKCEN", "CUSTOM")
   INFLATE w2, h2, 2
   ReDim ScaleMem(1 To w2, 1 To h2)
   CopyMemory ScaleMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   FillStruc bmscale, ScaleMem(), 2
   
   '===================================================
   ' 4 Special .arz for picLSlide
   '------------------------------
   N = 4
   BArray = LoadResData("WATCHITA", "CUSTOM")
   INFLATE w4, h4, 4
   ReDim WatchItMem(1 To w4, 1 To h4)
   CopyMemory WatchItMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   FillStruc bmwatchit, WatchItMem(), 4
   '------------------------------
   ' 4 Special .msz for picLSlide
   '------------------------------
   N = 41
   BArray = LoadResData("WATCHITM", "CUSTOM")
   INFLATE w4, h4, 41
   ReDim WatchItMaskMem(1 To w4, 1 To h4)
   CopyMemory WatchItMaskMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()

   '===================================================
   ' 5 Special .arz for picVScale
   '------------------------------
   N = 5
   BArray = LoadResData("PVSCALEA", "CUSTOM")
   INFLATE w5, h5, 5
   ReDim pvScaleMem(1 To w5, 1 To h5)
   CopyMemory pvScaleMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   FillStruc bmpvscale, pvScaleMem(), 5
   '------------------------------
   ' 5 Special .msk for picVScale
   '------------------------------
   N = 51
   BArray = LoadResData("PVSCALEM", "CUSTOM")
   INFLATE w5, h5, 51
   ReDim pvScaleMaskMem(1 To w5, 1 To h5)
   CopyMemory pvScaleMaskMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   
   '===================================================
   ' 6 Special .arz for picVSlider
   '------------------------------
   N = 6
   BArray = LoadResData("PVSLIDERA", "CUSTOM")
   INFLATE w6, h6, 6
   ReDim pvSliderMem(1 To w6, 1 To h6)
   CopyMemory pvSliderMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   FillStruc bmpvslider, pvSliderMem(), 6
   '------------------------------
   ' 6 Special .msz for picVSlider
   '------------------------------
   N = 61
   BArray = LoadResData("PVSLIDERM", "CUSTOM")
   INFLATE w6, h6, 61
   ReDim pvSliderMaskMem(1 To w6, 1 To h6)
   CopyMemory pvSliderMaskMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()

   '===================================================
   ' 7 Special .arz for picWheel
   '------------------------------
   N = 7
   BArray = LoadResData("TESTWHEEL", "CUSTOM")
   INFLATE w7, h7, 7
   ReDim WheelMem(1 To w7, 1 To h7)
   CopyMemory WheelMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   FillStruc bmwheel, WheelMem(), 7
   
   '===================================================
   ' 8 Special .arz for picCross UP
   '------------------------------
   N = 8
   BArray = LoadResData("CROSSUP", "CUSTOM")
   INFLATE w8, h8, 8
   ReDim CrossUpMem(1 To w8, 1 To h8)
   CopyMemory CrossUpMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   FillStruc bmcross, CrossUpMem(), 8
   
   '===================================================
   ' 8 Special .arz for picCross DN
   '------------------------------
   N = 81
   BArray = LoadResData("CROSSDN", "CUSTOM")
   INFLATE w8, h8, 81
   ReDim CrossDnMem(1 To w8, 1 To h8)
   CopyMemory CrossDnMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   
   '===================================================
   ' 9 Special .arz for picSwitch UP
   '------------------------------
   N = 9
   BArray = LoadResData("SWITCHUP2A", "CUSTOM")
   INFLATE w9, h9, 9
   ReDim SwitchUpMem(1 To w9, 1 To h9)
   CopyMemory SwitchUpMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   
   FillStruc bmswitch, SwitchUpMem(), 9
   '------------------------------
   ' 9 Special .msz for picSwitch  UP
   '------------------------------
   N = 91
   BArray = LoadResData("SWITCHUP2M", "CUSTOM")
   INFLATE w9, h9, 91
   ReDim SwitchUpMaskMem(1 To w9, 1 To h9)
   CopyMemory SwitchUpMaskMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()

   '===================================================
   ' 9 Special .arz for picSwitch DN
   '------------------------------
   BArray = LoadResData("SWITCHDN2A", "CUSTOM")
   INFLATE w9, h9, 92
   ReDim SwitchDnMem(1 To w9, 1 To h9)
   CopyMemory SwitchDnMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   '------------------------------
   ' 9 Special .msz for picSwitch  DN
   '------------------------------
   BArray = LoadResData("SWITCHDN2M", "CUSTOM")
   INFLATE w9, h9, 93
   ReDim SwitchDnMaskMem(1 To w9, 1 To h9)
   CopyMemory SwitchDnMaskMem(1, 1), uncomBArray(1), ByteLen2
   Erase uncomBArray()
   '===================================================
   
   Exit Sub
'================
FileErr:
   Close
   MsgBox "Error in LOAD_ARRAY, Num =" & Str$(N) & vbCr _
   & "w1=" & Str$(w1) & " h1=" & Str$(w2) & vbCr _
   & "ByteLen=" & Str$(ByteLen) & " ByteLen2= " & Str$(ByteLen2)
   
   aFileErrors = True
End Sub

Public Function zAtn2(ByVal Y As Single, ByVal X As Single) As Single
'0° to right, 0 to -pi#(-180°) anticlockwise, 0 to +pi#(+180°) clockwise
If X = 0 Then
    If Abs(Y) > Abs(X) Then   'Must be an overflow
        If Y > 0 Then zAtn2 = pi# / 2 Else zAtn2 = -pi# / 2
    Else
        zAtn2 = 0   'Must be an underflow
    End If
Else
    zAtn2 = Atn(Y / X)
    If (X < 0) Then
        If (Y < 0) Then zAtn2 = zAtn2 - pi# Else zAtn2 = zAtn2 + pi#
    End If
End If
End Function

Public Sub INFLATE(ByRef Wx As Long, ByRef Hy As Long, N As Long)
' In: compressed array BArray()
' Out: uncompressed array uncomBArray()
'      ByteLen2 = Wx * Hy * 4
'      Wx,Hy

'Public BArray() As Byte
'Public uncomBArray() As Byte
'Public ByteLen As Long
'Public ByteLen2 As Long
   
   ByteLen = UBound(BArray()) + 1
   ' Get W & H
   CopyMemory Wx, BArray(0), 4
   CopyMemory Hy, BArray(4), 4
   ' Cut off 8 Byte W & H
   'ByteLen = ByteLen - 8
   ReDim Preserve BArray(ByteLen)
   ' uncompressed size =
   ByteLen2 = Wx * Hy * 4
   ReDim uncomBArray(1 To ByteLen2)
   DoEvents
   
   ' **************************************
   Select Case uncompress( _
        uncomBArray(1), ByteLen2, _
        BArray(8), ByteLen)
   Case Z_MEM_ERROR
       MsgBox "Insufficient memory", vbExclamation, _
           "Compression Error" & Str$(N)
       Exit Sub
   Case Z_BUF_ERROR
       MsgBox "Buffer too small" & Str$(ByteLen2), vbExclamation, _
           "Compression Error" & Str$(N)
       Exit Sub
   Case Z_DATA_ERROR
       MsgBox "Input file corrupted", vbExclamation, _
           "Compression Error" & Str$(N)
       Exit Sub
   ' Else Z_OK.
   End Select
   Erase BArray

   ' **************************************
End Sub

