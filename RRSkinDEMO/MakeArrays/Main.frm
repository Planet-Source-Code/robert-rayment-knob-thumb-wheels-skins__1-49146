VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image Extraction"
   ClientHeight    =   5100
   ClientLeft      =   150
   ClientTop       =   105
   ClientWidth     =   3465
   DrawWidth       =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdACT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2.  Save Compressed Image Mask .msz"
      Height          =   390
      Index           =   2
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   3000
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Left            =   195
      TabIndex        =   3
      Top             =   150
      Width           =   3075
      Begin VB.PictureBox PIC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1725
         Left            =   60
         ScaleHeight     =   115
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   195
         TabIndex        =   4
         ToolTipText     =   "File = PIC"
         Top             =   45
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdACT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1.  Save Compressed Image Array .arz"
      Height          =   405
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2580
      Width           =   3015
   End
   Begin VB.CommandButton cmdACT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.  Load image"
      Height          =   420
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2010
      Width           =   3015
   End
   Begin VB.Label LabInfo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabInfo"
      Height          =   1200
      Left            =   255
      TabIndex        =   2
      Top             =   3690
      Width           =   3030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' True Color Image Extraction & Compression by Robert Rayment Oct 2003

' Compression routine uses zlib.dll available from
' www.winimage.com/zlibdll/
' Place in Windows/System or in same folder as app.
' For good example of usage see vb-helper.com
' also www.gzip.org/zlib/

' Image to Binary extraction
' NB. ArrMem() made from image as BGRA (W x H x 4 bytes/pixel)
'     ArrMem(W, H) -> compress into comBArray()
'     comBArray() saved as *.arz binary file
'     Default name same as image file

'     MaskMem() made from ArrMem() where black
'       pixels are set = -1 (ie 255,255,255,255, White, long value -1)
'       and the rest = 0
'     MaskMem() saved as *.msz binary file (W x H x 4 bytes/pixel)
'     Default name same as image file

'     For both put W & H at beginning of compressed data
'     NB Not at end since in EXE res data is 4 byte aligned
'        changing the end point of the data.
'      ReDim BArray(1 To ComSize + 8)
'      CopyMemory BArray(9), comBArray(1), ComSize
'      CopyMemory BArray(1), W, 4
'      CopyMemory BArray(5), H, 4
'     So the original size is available for uncompress
'     ie W & H giving original size = WxHx4
'      & W & H available for picture array.

Private Declare Function compress Lib "zlib.dll" _
   (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'Private Declare Function uncompress Lib "zlib.dll" _
'  (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
' zlib responses
Private Const Z_OK = 0
Private Const Z_DATA_ERROR = -3
Private Const Z_MEM_ERROR = -4
Private Const Z_BUF_ERROR = -5

Private CommonDialog1 As New OSDialog

' Files
Private PathSpec$, CurrentPath$
Private FileSpec$, FName$

Private ComSize As Long
Private ByteLen As Long
'

Private Sub Form_Load()
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrentPath$ = PathSpec$
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   W = 0 ' Picture Width
   H = 0 ' Picture Height
   
   ShowInfo
   
   aLoaded = False
   aArrSaved = False
   aMskSaved = False
   
   ReDim ArrMem(1 To 1)
   ReDim MaskMem(1 To 1)
End Sub

Private Sub ShowInfo()
a$ = vbCr
a$ = a$ & "  Width =" & Str$(W) & vbCr
a$ = a$ & "  Height =" & Str$(H) & vbCr
a$ = a$ & "  Array size =" & Str$((W * H * 4)) & " B" & vbCr
a$ = a$ & "  Compressed size =" & Str$(ComSize) & " B"
LabInfo = a$
End Sub

'#### FILE OPS ######################################

Private Sub cmdACT_Click(Index As Integer)
   
   Select Case Index
   Case 0   ' Load PIC
      Set CommonDialog1 = New OSDialog
      LOAD_PIC
      Set CommonDialog1 = Nothing
      PIC_ARRAY   ' Get DIBs
   
   Case 1   ' Save .arz
      If Not aLoaded Then
         MsgBox " No picture loaded yet", vbInformation, "Save Array .arz"
         Exit Sub
      End If
      
      Set CommonDialog1 = New OSDialog
      SAVE_ARRAY Index
      Set CommonDialog1 = Nothing
   
   Case 2   ' Save .msz
      If Not aLoaded Then
         MsgBox " No picture loaded yet", vbInformation, "Save Mask .msz"
         Exit Sub
      End If
      
      If Not aArrSaved Then
         MsgBox " Save *.arz first", vbInformation, "Save Mask .msk"
         Exit Sub
      End If
      
      Set CommonDialog1 = New OSDialog
      SAVE_MASK Index
      Set CommonDialog1 = Nothing
      
   End Select
End Sub

Private Sub LOAD_PIC()
Dim Title$, Filt$, InDir$
Dim pnum As Integer
Dim pnum2 As Integer

   MousePointer = vbDefault
   
   Title$ = "Load a picture file"
   Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
   InDir$ = CurrentPath$
   
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
   If Len(FileSpec$) = 0 Then
      Close
      Exit Sub
   End If
   
   pnum = InStrRev(FileSpec$, "\")
   CurrentPath$ = Left$(FileSpec$, pnum)
   pnum2 = InStrRev(FileSpec$, ".")
   FName$ = Mid$(FileSpec$, pnum + 1, pnum2 - pnum - 1)
   
   PIC.Picture = LoadPicture
   PIC.Picture = LoadPicture(FileSpec$)
   PIC.ToolTipText = " " & FileSpec$ & " "
   
   W = PIC.Width \ STX
   H = PIC.Height \ STY
   
   ShowInfo
   
   aLoaded = True
   aArrSaved = False
   aMskSaved = False
   
   PIC_ARRAY
   
   DoEvents
End Sub

Private Sub PIC_ARRAY()
   ReDim ArrMem(1 To W, 1 To H)
   FillBMPStruc
   GETDIB PIC.Image
End Sub

Private Sub SAVE_ARRAY(Index As Integer)
Dim FileSpec$, Title$, Filt$, InDir$
   
   Title$ = "Save compressed binary File"
   Filt$ = "Save arz (.arz)|*.arz"
   InDir$ = CurrentPath$
   FileSpec$ = FName$
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd

   If Len(FileSpec$) = 0 Then
      aArrSaved = False
      Exit Sub
   Else
      
      FixExtension FileSpec$, "arz"
      ' .arz Image
      ByteLen = 4 * W * H
      ReDim BArray(1 To ByteLen)
      CopyMemory BArray(1), ArrMem(1, 1), ByteLen
         
      If Len(Dir$(FileSpec$)) <> 0 Then Kill FileSpec$
      
      ' **************************************
      ' Compress.
   
      ' Allocate the smallest allowed compression
      ' buffer (1% larger than the uncompressed data
      ' plus 12 bytes).
      ComSize = 1.01 * ByteLen + 12
      ReDim comBArray(1 To ComSize)
   
       ' Compress the bytes.
      Select Case compress( _
          comBArray(1), ComSize, _
          BArray(1), ByteLen)
      Case Z_MEM_ERROR
          MsgBox "Insufficient memory", vbExclamation, _
              "Compression Error"
          Exit Sub
      Case Z_BUF_ERROR
          MsgBox "Buffer too small", vbExclamation, _
              "Compression Error"
          Exit Sub
      ' Else Z_OK.
      End Select
   
      ' Shrink the compressed buffer to fit.
      ' & Expand comBArray to take original W & H
      'ReDim Preserve comBArray(1 To ComSize + 8)
      ReDim BArray(1 To ComSize + 8)
      CopyMemory BArray(9), comBArray(1), ComSize
      CopyMemory BArray(1), W, 4
      CopyMemory BArray(5), H, 4
      ' **************************************
         
      Open FileSpec$ For Binary As #1
      Put #1, , BArray() 'ArrMem()
      Close
      
      ComSize = FileLen(FileSpec$)
      
      aArrSaved = True
      
      Erase BArray(), comBArray()
   
      ShowInfo
   End If
End Sub


Private Sub SAVE_MASK(Index As Integer)
Dim FileSpec$, Title$, Filt$, InDir$

   Title$ = "Save compressed binary Mask File"
   Filt$ = "Save msz (.msz)|*.msz"
   InDir$ = CurrentPath$
   FileSpec$ = FName$
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd

   If Len(FileSpec$) = 0 Then
      aMskSaved = False
      Exit Sub
   Else
      
      ' Make mask Black->White, Non-black to black
      ReDim MaskMem(1 To W, 1 To H)
      
      For j = 1 To H
      For i = 1 To W
         If ArrMem(i, j) = 0 Then
            MaskMem(i, j) = -1   ' White (255,255,255,255)
         Else
            MaskMem(i, j) = 0
         End If
      Next i
      Next j
         
      FixExtension FileSpec$, "msz"
      
      If Len(Dir$(FileSpec$)) <> 0 Then Kill FileSpec$
      
      ByteLen = 4 * W * H
      ReDim BArray(1 To ByteLen)
      CopyMemory BArray(1), MaskMem(1, 1), ByteLen
      
      ' **************************************
      ' Compress.
   
      ' Allocate the smallest allowed compression
      ' buffer (1% larger than the uncompressed data
      ' plus 12 bytes).
      ComSize = 1.01 * ByteLen + 12
      ReDim comBArray(1 To ComSize)
   
       ' Compress the bytes.
      Select Case compress( _
          comBArray(1), ComSize, _
          BArray(1), ByteLen)
      Case Z_MEM_ERROR
          MsgBox "Insufficient memory", vbExclamation, _
              "Compression Error"
          Exit Sub
      Case Z_BUF_ERROR
          MsgBox "Buffer too small", vbExclamation, _
              "Compression Error"
          Exit Sub
      ' Else Z_OK.
      End Select
   
      ' Shrink the compressed buffer to fit.
      ' & Expand comBArray to take original W & H
      ReDim BArray(1 To ComSize + 8)
      CopyMemory BArray(9), comBArray(1), ComSize
      CopyMemory BArray(1), W, 4
      CopyMemory BArray(5), H, 4
    ' **************************************
      
      Open FileSpec$ For Binary As #1
      Put #1, , BArray() 'ArrMem()
      Close
      
      ComSize = FileLen(FileSpec$)
      
      aArrSaved = True
      
      Erase BArray(), comBArray()
      
      aMskSaved = True
      
      ShowInfo
   End If
End Sub
'#### END FILE OPS ######################################

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'aLoaded = True/False
'aArrSaved = True/False
   If UnloadMode = 0 Then    'Close on Form1 pressed
      If aLoaded = True Then
         resp = MsgBox("Quit", vbQuestion + vbYesNo, "Image EXtraction")
         If resp = vbYes Then
            Cancel = 0
         Else
            Cancel = 1
         End If
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   End
End Sub
