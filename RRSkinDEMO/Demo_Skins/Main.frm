VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "      BA"
   ClientHeight    =   3840
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   3840
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSwitch2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1170
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   14
      Top             =   105
      Width           =   240
   End
   Begin VB.PictureBox picSwitch 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1455
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   13
      Top             =   105
      Width           =   240
   End
   Begin VB.PictureBox picCross 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   1740
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   11
      Top             =   105
      Width           =   360
   End
   Begin VB.PictureBox picWheel 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3255
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   10
      Top             =   2850
      Width           =   285
   End
   Begin VB.PictureBox picVSlider 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3285
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   8
      Top             =   1245
      Width           =   270
   End
   Begin VB.PictureBox picVScale 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   3330
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   7
      Top             =   525
      Width           =   120
   End
   Begin VB.PictureBox picLSide 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   -945
      MousePointer    =   9  'Size W E
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   6
      Top             =   270
      Width           =   1200
   End
   Begin VB.PictureBox picSlider 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1800
      MouseIcon       =   "Main.frx":058A
      MousePointer    =   99  'Custom
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   5
      Top             =   915
      Width           =   180
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C000&
      Height          =   165
      Left            =   1620
      Picture         =   "Main.frx":0E54
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   3
      Top             =   1785
      Width           =   180
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1605
      MaxLength       =   4
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1185
      Width           =   585
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2955
      Width           =   435
   End
   Begin VB.PictureBox picScale 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   210
      Left            =   900
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   1
      Top             =   900
      Width           =   2010
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   11
      Left            =   480
      Picture         =   "Main.frx":1D36
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   10
      Left            =   480
      Picture         =   "Main.frx":2078
      Stretch         =   -1  'True
      Top             =   2310
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   9
      Left            =   480
      Picture         =   "Main.frx":23BA
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   8
      Left            =   480
      Picture         =   "Main.frx":26FC
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   7
      Left            =   480
      Picture         =   "Main.frx":2A3E
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   6
      Left            =   480
      Picture         =   "Main.frx":2D80
      Stretch         =   -1  'True
      Top             =   1230
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   5
      Left            =   180
      Picture         =   "Main.frx":30C2
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   4
      Left            =   180
      Picture         =   "Main.frx":3404
      Stretch         =   -1  'True
      Top             =   2310
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   3
      Left            =   180
      Picture         =   "Main.frx":3746
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   2
      Left            =   180
      Picture         =   "Main.frx":3A88
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   1
      Left            =   180
      Picture         =   "Main.frx":3DCA
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   240
   End
   Begin VB.Image imgSkin 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   0
      Left            =   180
      Picture         =   "Main.frx":410C
      Stretch         =   -1  'True
      Top             =   1230
      Width           =   240
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   300
      Left            =   3435
      Shape           =   4  'Rounded Rectangle
      Top             =   3495
      Width           =   360
   End
   Begin VB.Image imgCtrl 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   3
      Left            =   3075
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Main.frx":444E
      Stretch         =   -1  'True
      ToolTipText     =   "Exit "
      Top             =   105
      Width           =   240
   End
   Begin VB.Image imgCtrl 
      Height          =   180
      Index           =   2
      Left            =   2745
      MousePointer    =   10  'Up Arrow
      Picture         =   "Main.frx":4530
      Stretch         =   -1  'True
      ToolTipText     =   "Small "
      Top             =   90
      Width           =   240
   End
   Begin VB.Image imgCtrl 
      Height          =   180
      Index           =   1
      Left            =   2460
      MousePointer    =   10  'Up Arrow
      Picture         =   "Main.frx":4612
      Stretch         =   -1  'True
      ToolTipText     =   "Large "
      Top             =   75
      Width           =   240
   End
   Begin VB.Image imgCtrl 
      BorderStyle     =   1  'Fixed Single
      Height          =   180
      Index           =   0
      Left            =   2175
      MousePointer    =   15  'Size All
      Picture         =   "Main.frx":46F4
      Stretch         =   -1  'True
      ToolTipText     =   "Move "
      Top             =   90
      Width           =   240
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   165
      Top             =   2880
      Width           =   600
   End
   Begin VB.Image imgLight 
      Height          =   120
      Left            =   390
      Picture         =   "Main.frx":47D6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   180
      Left            =   3390
      TabIndex        =   12
      Top             =   135
      Width           =   270
   End
   Begin VB.Label LabVScale 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   3315
      TabIndex        =   9
      Top             =   2625
      Width           =   180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderColor     =   &H00FFFFFF&
      Height          =   2730
      Left            =   3165
      Shape           =   4  'Rounded Rectangle
      Top             =   510
      Width           =   465
   End
   Begin VB.Label LabName 
      Caption         =   " Demo by Robert Rayment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   3390
      Width           =   2190
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   645
      Left            =   3210
      Shape           =   4  'Rounded Rectangle
      Top             =   2565
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1  (Main.frm)

' Blittin' Madness  by Robert Rayment

Option Explicit

' For resizing all controls
Private Type CSizes
  xL As Single
  yT As Single
  xW As Single
  yH As Single
End Type
Dim SizeArr() As CSizes

' For moving objects
Dim Xfra As Single
Dim Yfra As Single

' Initial font sizes
'Dim chkFontSizeO  As Long  ' Check box  Check1
Dim txFontSizeO As Long ' Text box Text1
Dim LabFontSizeO As Long   ' LabName
'Dim picLSideFontSizeO As Long ' picLSide

Dim FWO As Long, FHO As Long  ' Form Fixed small size

Dim SliderValue As Integer
Dim VSliderValue As Integer
Dim prevVSliderValue As Integer
Dim iPicture1 As Long

Dim SWidth As Long
Dim SHeight As Long
Dim aFirstLoad As Boolean

Dim aFileErrors As Boolean

' Gen back color
Dim BCul As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_APPWORKSPACE = 12
Dim mAppWorkSpace As Long





Private Sub Form_Load()
   CurrentPath$ = App.Path
   If Right$(CurrentPath$, 1) <> "\" Then CurrentPath$ = CurrentPath$ & "\"
   PathSpec$ = CurrentPath$ & "ImageArrays\"
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   aResize = False
   aChecked = False
   FWO = 256
   FHO = 256
   
   Me.Width = FWO * STX
   Me.Height = FHO * STY
   
   mAppWorkSpace = GetSysColor(COLOR_APPWORKSPACE)
   Me.BackColor = mAppWorkSpace
   
   W = 0 ' Form Width
   H = 0 ' Form Height
   
   aLoaded = False
   ReDim ArrMem(1 To 1)
   
   ' Get starting values
   XMag = 1
   YMag = 1
   xmagslide = 1
   ymagslide = 1
   
   ORGFW = Me.Width
   ORGFH = Me.Height
   
   ' Position imgCtrls
   imgCtrl(0).Left = 150
   imgCtrl(1).Left = 168
   imgCtrl(2).Left = 187
   imgCtrl(3).Left = 205
   For i = 0 To 3
      With imgCtrl(i)
         .Top = 9
         .Width = 16
         .Height = 12
      End With
   Next i
   
   LOAD_ARRAYS aFileErrors
   
   If aFileErrors Then
      MsgBox " Load from RES Error", vbExclamation, " Loading Arrays"
      Form_Unload 0
      Exit Sub
   End If
   
   Me.Show
   
   ' Initial font sizes
   'chkFontSizeO = Check1.FontSize
   txFontSizeO = Text1.FontSize
   LabFontSizeO = LabName.FontSize
   'picLSideFontSizeO = picLSide.FontSize
   
   ' Center sliders
   picSlider.Left = picScale.Left + picScale.Width / 2 - picSlider.Width / 2
   SliderValue = 50
   Text1.Text = Str$(SliderValue)
   
   picVSlider.Top = picVScale.Top + picVScale.Height / 2 - picVSlider.Height / 2
   VSliderValue = 31
   LabVScale = Str$(VSliderValue)
   
   InitSizeArr
      
   ' Identify Picture1 place in Controls()
   For i = 0 To Controls.Count - 1
    If Controls(i).Name = "Picture1" Then
       iPicture1 = i
       Exit For
    End If
   Next i
   
   aResize = True
   KeyPreview = True  ' To allow Esc key to Quit
   
   ' Set hand cursors
   ' This is B/W in IDE but colored in EXE.
   picSlider.MousePointer = vbCustom
   picSlider.MouseIcon = LoadResPicture(101, vbResCursor)
   
   picVSlider.MousePointer = vbCustom
   picVSlider.MouseIcon = LoadResPicture(101, vbResCursor)
   
   picWheel.MousePointer = vbCustom
   picWheel.MouseIcon = LoadResPicture(101, vbResCursor)
   
   picCross.MousePointer = vbCustom
   picCross.MouseIcon = LoadResPicture(101, vbResCursor)
   
   picSwitch.MousePointer = vbCustom
   picSwitch.MouseIcon = LoadResPicture(101, vbResCursor)
   
   picSwitch2.MousePointer = vbCustom
   picSwitch2.MouseIcon = LoadResPicture(101, vbResCursor)
   
   ' Load sound data
   'WAVSWITCH = LoadResData("SWITCH", "CUSTOM")
   WAVSWITCH = StrConv(LoadResData("SWITCH", "CUSTOM"), vbUnicode)
   WAVSLIDE = StrConv(LoadResData("SLIDE", "CUSTOM"), vbUnicode)
   WAVTIC = StrConv(LoadResData("TIC", "CUSTOM"), vbUnicode)
   WAVTIC2 = StrConv(LoadResData("TIC2", "CUSTOM"), vbUnicode)

   ' Switch
   SwitchYPrev = 0
   aSwitch2 = False
   imgLight.Visible = False

'=============================================

   ' Tile pic
   With Picture1
      .Width = 2 + 0.8 * 100 * 3.6
      .Height = 2 + 0.95 * 100 * 2.7
   End With
   TILEPIC
   ResizePicture1
'=============================================
   
   aTile = False
   aFirstLoad = True
   
   imgSkin_Click 0   ' BLUE skin

End Sub

Private Sub TILEPIC()
   Dim ehh As Long
   Dim eww As Long
   
   'Picture1.Cls
   'Picture1.Picture = LoadPicture(CurrentPath$ & "Weave.bmp")
   'Picture1.Picture = LoadPicture(CurrentPath$ & "Mesh.bmp")
   Picture1.Refresh
   
   ehh = 36
   eww = 34
   'ehh = 32
   'eww = 32
   
   
   For j = 0 To Picture1.Height Step ehh
   For i = 0 To Picture1.Width Step eww
      BitBlt Picture1.hdc, i, j, eww, ehh, Picture1.hdc, 0, 0, vbSrcCopy
   Next i
   Next j
   Picture1.Refresh
End Sub

Private Sub TILEFORM(N As Integer)
   
   Dim ehh As Long
   Dim eww As Long
   
   Me.Picture = imgSkin(N).Picture
   
   ehh = 16
   eww = 16
   
   ' Make 64 x 64 tile
   For j = 0 To 2 * ehh Step ehh
   For i = 0 To 2 * eww Step eww
      BitBlt Me.hdc, i, j, eww, ehh, Me.hdc, 0, 0, vbSrcCopy
   Next i
   Next j
   
   ' Tile 64 x 64
   For j = 0 To Me.Width \ 2 - 1 Step 2 * ehh
   For i = 0 To Me.Height \ 2 - 1 Step 2 * eww
      BitBlt Me.hdc, i, j, 2 * eww, 2 * ehh, Me.hdc, 0, 0, vbSrcCopy
   Next i
   Next j
   
   
   
   Me.Refresh
End Sub

Private Sub imgSkin_Click(Index As Integer)

   Screen.MousePointer = vbHourglass
   
   Select Case Index
   Case 0   ' BLUE
      ' Get compressed array from resource file
      BArray = LoadResData("BLUE", "CUSTOM")
      aTile = False
   Case 1   ' PINK
      ' Get compressed array from resource file
      BArray = LoadResData("PINK", "CUSTOM")
      aTile = False
   
   Case Else
      TILEFORM Index
      aTile = True
   End Select

   If Not aTile Then
   
      INFLATE W, H, 0
      
      ReDim ArrMem(1 To W, 1 To H)
      CopyMemory ArrMem(1, 1), uncomBArray(1), ByteLen2
      
      Erase uncomBArray()
      
      FillStruc bm, ArrMem(), 0
      
      DoEvents
      aLoaded = True
   End If
'----------------------------------------------------------
   If aFirstLoad Then
      ' Reduce to skin size first
      SWidth = GetSystemMetrics(SM_CXSCREEN) * STX
      SHeight = GetSystemMetrics(SM_CYSCREEN) * STY
      With Me
         .Cls
         .Width = W * STX
         .Height = H * STY
         'Centre form
         .Top = (SHeight - .Height) / 2
         .Left = (SWidth - .Width) / 2
         ORGFW = .Width
         ORGFH = .Height
      End With

      XMag = 1
      YMag = 1
      xmagslide = 1
      ymagslide = 1

      aFirstLoad = False
   End If
'----------------------------------------------------------
   
   ResizePicture1
   
   If aChecked Then
      SWR hWnd, CRR(6, 6, (FWO * XMag), (FHO * YMag), 50, 50), True
   Else
      SWR hWnd, CRR(0, 0, (FWO * XMag), (FHO * YMag), 0, 0), True
   End If

   ' Load in new skin
   DoEvents
   If Not aTile Then ShowWholePicture
   ResizeControls
   MaskSlider
   MaskWatchIt
   MaskVSlider
   MaskSwitchUP
   MaskSwitch2UP

   BCul = Me.Point((Check1.Left + Check1.Width) + 6, Check1.Top + 10)
   Check1.BackColor = BCul
   BCul = Me.Point((LabName.Left + LabName.Width) + 6, LabName.Top + 10)
   LabName.BackColor = BCul

   VSliderValue = 31
   LabVScale = Str$(VSliderValue)

   StretchDIBits picWheel.hdc, _
   0, 0, picWheel.Width, picWheel.Height, _
   VSliderValue * 19, 0, 19, 19, _
   WheelMem(1, 1), bmwheel, _
   DIB_RGB_COLORS, _
   vbSrcCopy

   picWheel.Refresh

   StretchDIBits picCross.hdc, _
   0, 0, picCross.Width, picCross.Height, _
   0, 0, w8, w8, _
   CrossUpMem(1, 1), bmcross, _
   DIB_RGB_COLORS, _
   vbSrcCopy

   picCross.Refresh

   Picture1.SetFocus
   Screen.MousePointer = vbDefault
End Sub


Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Rounded rectangle check
   aChecked = Not aChecked
   Check1.Value = 0
   
   If aChecked Then
      SWR hWnd, CRR(6, 6, (FWO * XMag), (FHO * YMag), 50, 50), True
      Shape3.Shape = 4
   Else
      SWR hWnd, CRR(0, 0, (FWO * XMag), (FHO * YMag), 0, 0), True
      Shape3.Shape = 0
   End If
   
   Picture1.SetFocus
   
   DoEvents
   
End Sub



'#### TR BUTTONS ##################################################
Private Sub imgCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Top ctrls Move, Max, Normal, Exit
   SWidth = GetSystemMetrics(SM_CXSCREEN) * STX
   SHeight = GetSystemMetrics(SM_CYSCREEN) * STY
   
   Select Case Index
   Case 0   ' Move
      Xfra = X: Yfra = Y
   Case 1   ' Large
      With Me
         .Left = 0.1 * SWidth
         .Top = 0.1 * SHeight
         .Width = 0.9 * SWidth
         .Height = 0.9 * SHeight
         XMag = .Width / ORGFW
         YMag = .Height / ORGFH
         xmagslide = (.Width / STX) / FWO
         ymagslide = (.Height / STY) / FHO
         .Top = (SHeight - .Height) / 2
         .Left = (SWidth - .Width) / 2
      End With
      
      If aChecked Then
         SWR hWnd, CRR(6, 6, (FWO * XMag), (FHO * YMag), 50, 50), True
      Else
         SWR hWnd, CRR(0, 0, (FWO * XMag), (FHO * YMag), 0, 0), True
      End If
      
      ResizePicture1
      
      Xfra = 0: Yfra = 0
      
      If Not aTile Then ShowWholePicture
      ResizeControls
      MaskSlider
      MaskWatchIt
      MaskVSlider
      
      picCross_MouseUp 1, 0, 0, 0
      MaskSwitchUP
      MaskSwitch2UP
      
      'Centre form
      With Me
         .Top = (SHeight - .Height) / 2
         .Left = (SWidth - .Width) / 2
      End With
   
      VSliderValue = 31
      LabVScale = LTrim$(Str$(VSliderValue))
      
      StretchDIBits picWheel.hdc, _
      0, 0, picWheel.Width, picWheel.Height, _
      VSliderValue * 19, 0, 19, 19, _
      WheelMem(1, 1), bmwheel, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      picWheel.Refresh
   
   Case 2   ' Small
      With Me
         .Width = ORGFW: .Height = ORGFH
         'Centre form
         .Top = (SHeight - .Height) / 2
         .Left = (SWidth - .Width) / 2
      End With
      Xfra = 0: Yfra = 0
      XMag = 1: YMag = 1
      xmagslide = 1
      ymagslide = 1
      
      If aChecked Then
         SWR hWnd, CRR(6, 6, (FWO * XMag), (FHO * YMag), 50, 50), True
      Else
         SWR hWnd, CRR(0, 0, (FWO * XMag), (FHO * YMag), 0, 0), True
      End If
      
      ResizePicture1
      
      If Not aTile Then ShowWholePicture
      ResizeControls
      MaskSlider
      MaskWatchIt
      MaskVSlider
      
      MaskSwitchUP
      MaskSwitch2UP
      picCross_MouseUp 1, 0, 0, 0
      
      VSliderValue = 31
      LabVScale = LTrim$(Str$(VSliderValue))
      
      StretchDIBits picWheel.hdc, _
      0, 0, picWheel.Width, picWheel.Height, _
      VSliderValue * 19, 0, 19, 19, _
      WheelMem(1, 1), bmwheel, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      picWheel.Refresh
   
   Case 3   ' Exit
      Form_Unload 0
   End Select
End Sub

Private Sub imgCtrl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
   Case 0   ' Move
      If Button = vbLeftButton Then
            Me.Left = Me.Left + (X - Xfra)
            Me.Top = Me.Top + (Y - Yfra)
      End If
   End Select
End Sub
'#### END TR BUTTONS ##################################################


'### SWITCHES ######################################################

Private Sub picSwitch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      MaskSwitchDN
      
      MaskSwitch2DN
      
      sndPlaySound WAVSWITCH, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
      
      With Picture1
         .Width = 2 + 0.8 * 100 * XMag
         .Height = 2 + 0.95 * 100 * YMag
         .Refresh
      End With
   
   End If
End Sub

Private Sub picSwitch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      MaskSwitchUP
      
      MaskSwitch2UP
      
      sndPlaySound WAVSWITCH, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
      
      With Picture1
         .Width = 2 + 0.8 * SliderValue * XMag
         .Height = 2 + 0.95 * SliderValue * YMag
         .Refresh
      End With
   End If
End Sub

Private Sub picSwitch2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      If Y < SwitchYPrev Then
         ' Switch up
         MaskSwitch2UP
         SwitchYPrev = Y
         If aSwitch2 Then
            sndPlaySound WAVSWITCH, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
            aSwitch2 = False
         End If
      Else
         ' Switch dn
         MaskSwitch2DN
         SwitchYPrev = Y
         If Not aSwitch2 Then
            sndPlaySound WAVSWITCH, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
            aSwitch2 = True
         End If
      
      End If
   
   End If
End Sub

Private Sub picSwitch2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   sndPlaySound "", SND_PURGE
End Sub

'### END SWITCHES ######################################################

'#### CROSS ##################################################
Private Sub picCross_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      StretchDIBits picCross.hdc, _
      0, 0, picCross.Width, picCross.Height, _
      0, 0, w8, w8, _
      CrossDnMem(1, 1), bmcross, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      picCross.Refresh
      
      With Picture1
         .Width = 2 + 0.8 * 100 * XMag
         .Height = 2 + 0.95 * 100 * YMag
         .Refresh
      End With
   End If
End Sub

Private Sub picCross_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      StretchDIBits picCross.hdc, _
      0, 0, picCross.Width, picCross.Height, _
      0, 0, w8, w8, _
      CrossUpMem(1, 1), bmcross, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      picCross.Refresh
      
      With Picture1
         .Width = 2 + 0.8 * SliderValue * XMag
         .Height = 2 + 0.95 * SliderValue * YMag
         .Refresh
      End With
   
   End If
End Sub
'#### END CROSS ##################################################


'### VSLIDER & WHEEL ##############################################

Private Sub picVScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PTop As Long
   If Button = vbLeftButton Then
      PTop = Y + picVScale.Top - picVSlider.Height / 2
      If PTop < picVScale.Top Then
         PTop = picVScale.Top
      ElseIf PTop > picVScale.Top + picVScale.Height - picVSlider.Height + 2 Then
         PTop = picVScale.Top + picVScale.Height - picVSlider.Height + 2
      End If
      picVSlider.Top = PTop
      
      MaskJustVSlider
      
      VSliderValue = (PTop - picVScale.Top) * 63 / ((picVScale.Height - picVSlider.Height) * 1.02)
      ' Fiddle factors
      If VSliderValue < 0 Then VSliderValue = 0
      If VSliderValue >= 63 Then VSliderValue = 63
      LabVScale = LTrim$(Str$(VSliderValue))
   
      StretchDIBits picWheel.hdc, _
      0, 0, picWheel.Width, picWheel.Height, _
      VSliderValue * 19, 0, 19, 19, _
      WheelMem(1, 1), bmwheel, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      picWheel.Refresh
   
      sndPlaySound WAVTIC2, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_MEMORY
   
   End If

End Sub

Private Sub picVSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xfra = X
   Yfra = Y
End Sub

Private Sub picVSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PTop As Long
   If Button = vbLeftButton Then
      PTop = picVSlider.Top + (Y - Yfra)
      If PTop < picVScale.Top Then
         PTop = picVScale.Top
      ElseIf PTop > picVScale.Top + picVScale.Height - picVSlider.Height + 2 Then
         PTop = picVScale.Top + picVScale.Height - picVSlider.Height + 2
      End If
      picVSlider.Top = PTop
      
      MaskJustVSlider
      
      VSliderValue = (PTop - picVScale.Top) * 63 / ((picVScale.Height - picVSlider.Height) * 1.02)
      ' Fiddle factors
      If VSliderValue < 0 Then VSliderValue = 0
      If VSliderValue >= 63 Then VSliderValue = 63
      LabVScale = LTrim$(Str$(VSliderValue))
   
      StretchDIBits picWheel.hdc, _
      0, 0, picWheel.Width, picWheel.Height, _
      VSliderValue * 19, 0, 19, 19, _
      WheelMem(1, 1), bmwheel, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      picWheel.Refresh
      If VSliderValue <> prevVSliderValue Then
         sndPlaySound WAVTIC2, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_MEMORY
         prevVSliderValue = VSliderValue
      End If
   End If

End Sub

Private Sub picVSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   sndPlaySound "", SND_PURGE
   LabVScale = LTrim$(Str$(VSliderValue))
End Sub

Private Sub picWheel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PTop As Long
'zAngle = zAtn2(DY, DX)
   If Button = vbLeftButton Then
      
      xcw = picWheel.Width / 2
      ycw = picWheel.Height / 2
      zAngle = zAtn2(Y - ycw, X - xcw)
      If zAngle > zprevAngle Then   ' Clockwise +
         If VSliderValue < 63 Then VSliderValue = VSliderValue + 1
      Else     ' Anti-clockwise
         If VSliderValue > 0 Then VSliderValue = VSliderValue - 1
      End If
   
      LabVScale = LTrim$(Str$(VSliderValue))
   
      StretchDIBits picWheel.hdc, _
      0, 0, picWheel.Width, picWheel.Height, _
      VSliderValue * 19, 0, 19, 19, _
      WheelMem(1, 1), bmwheel, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      picWheel.Refresh
      zprevAngle = zAngle
      
      sndPlaySound WAVTIC2, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_MEMORY
      
   End If
   

'      'VSliderValue = (PTop - picVScale.Top) * 63 / ((picVScale.Height - picVSlider.Height) * 1.02)
'      PTop = VSliderValue * ((picVScale.Height - picVSlider.Height) * 1.02) / 63 + picVScale.Top
'
'      If PTop < picVScale.Top Then
'         PTop = picVScale.Top
'      ElseIf PTop > picVScale.Top + picVScale.Height - picVSlider.Height + 2 Then
'         PTop = picVScale.Top + picVScale.Height - picVSlider.Height + 2
'      End If
'      picVSlider.Top = PTop
'      picVScale.Refresh
'      'Me.Refresh
'      picVSlider.Refresh
'      DoEvents
'
'      Sleep 1000
'   End If
End Sub

Private Sub picWheel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PTop As Long
   If Button = vbLeftButton Or Button = vbRightButton Then
   
      PTop = VSliderValue * ((picVScale.Height - picVSlider.Height) * 1.02) / 63 + picVScale.Top
      
      If PTop < picVScale.Top Then
         PTop = picVScale.Top
      ElseIf PTop > picVScale.Top + picVScale.Height - picVSlider.Height + 2 Then
         PTop = picVScale.Top + picVScale.Height - picVSlider.Height + 2
      End If
      picVSlider.Top = PTop
      
      MaskJustVSlider

      sndPlaySound WAVTIC2, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_MEMORY
      sndPlaySound "", SND_PURGE
   
   End If
End Sub
'### END VSLIDER & WHEEL ##############################################


'#### SIDE WINDOW ###############################################

Private Sub picLSide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xfra = X
   Yfra = Y
End Sub

Private Sub picLSide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PLeft As Long
   If Button = vbLeftButton Then
      PLeft = picLSide.Left + (X - Xfra)
      If PLeft + picLSide.Width < 4 * XMag Then PLeft = 4 * XMag - picLSide.Width
      If PLeft > 4 * XMag Then PLeft = 4 * XMag
      picLSide.Left = PLeft
      MaskWatchIt
   End If
End Sub

Private Sub picLSide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picLSide.Left = -picLSide.Width + 12 * XMag
   MaskWatchIt
      
   sndPlaySound WAVSLIDE, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_MEMORY
End Sub
'#### END SIDE WINDOW ###############################################

'#### PIC SLIDER ###########################

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim X As Single
Dim zR As Single

If KeyAscii = 13 Then
   KeyAscii = 0   ' Stop Beep
   
   If Not IsNumeric(Text1.Text) Then
      Text1.Text = Trim$(Str$(SliderValue))
      picScale.SetFocus
      Exit Sub
   End If
   
   SliderValue = Val(Text1.Text)
   
   If SliderValue < 0 Then
      SliderValue = 0
      Text1.Text = "0"
   End If
   If SliderValue > 100 Then
      SliderValue = 100
      Text1.Text = "100"
   End If
   
   zR = 100 / (picScale.Width - picSlider.Width)
   X = SliderValue / zR + picSlider.Width / 2
   picScale.SetFocus
   picScale_MouseDown 1, 0, X, 0
End If
End Sub

Private Sub picScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xfra = 0
   Yfra = 0
   picSlider.Left = picScale.Left + X - picSlider.Width / 2
   picSlider_MouseMove 1, 0, 0, 0
End Sub

Private Sub picSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xfra = X
   Yfra = Y
End Sub

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PLeft As Long
Dim zR As Single
   
   If Button = vbLeftButton Then
      
      
      PLeft = picSlider.Left + (X - Xfra)
      If PLeft < picScale.Left Then
         PLeft = picScale.Left
      ElseIf PLeft > picScale.Left + picScale.Width - picSlider.Width Then
         PLeft = picScale.Left + picScale.Width - picSlider.Width
      End If
      picSlider.Left = PLeft
      
      MaskSlider
      
      zR = 100 / (picScale.Width - picSlider.Width)
      SliderValue = (picSlider.Left - picScale.Left) * zR
      Text1.Text = Str$(SliderValue)
      
      With Picture1
         .Width = 2 + 0.8 * SliderValue * XMag
         .Height = 2 + 0.95 * SliderValue * YMag
         .Refresh
         SizeArr(iPicture1).xW = .Width
         SizeArr(iPicture1).yH = .Height
      End With
      
      sndPlaySound WAVTIC, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_MEMORY
      'sndPlaySound WAVTIC, SND_SYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_MEMORY
   
   End If
End Sub

Private Sub picSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   sndPlaySound "", SND_PURGE
End Sub
'#### END PIC SLIDER ###########################

Private Sub ResizePicture1()
   With Picture1
      .Width = 10
      .Height = 10
      .Refresh
      SizeArr(iPicture1).xW = .Width
      SizeArr(iPicture1).yH = .Height
   End With
End Sub

Private Sub InitSizeArr()
On Error Resume Next
   ReDim SizeArr(0 To Controls.Count - 1)
   For i = 0 To Controls.Count - 1
      SizeArr(i).xL = Controls(i).Left
      SizeArr(i).yT = Controls(i).Top
      SizeArr(i).xW = Controls(i).Width
      SizeArr(i).yH = Controls(i).Height
   Next i
End Sub

Private Sub ResizeControls()
On Error Resume Next
   For i = 0 To Controls.Count - 1
'LabFontSizeO
      ' Do font resizing first !!
      If TypeOf Controls(i) Is Label Then
         If XMag < YMag Then
            Controls(i).FontSize = LabFontSizeO * XMag
         Else
            Controls(i).FontSize = LabFontSizeO * YMag
         End If
      End If
      
      If TypeOf Controls(i) Is TextBox Then
         ' Both TextBoxes
         Controls(i).FontSize = txFontSizeO * (XMag + YMag) / 2
      End If
      
      'If TypeOf Controls(i) Is CheckBox Then
      '   Controls(i).FontSize = chkFontSizeO * (XMag + YMag) / 2
      'End If
      
'      If Controls(i).Name = "picLSide" Then
'         Controls(i).FontSize = picLSideFontSizeO * (XMag + YMag) / 2
'      End If

      Controls(i).Move _
      SizeArr(i).xL * XMag, _
      SizeArr(i).yT * YMag, _
      SizeArr(i).xW * XMag, _
      SizeArr(i).yH * YMag
      Controls(i).Refresh
   
   Next i
   
   ShowSlider
   
   SWR picCross.hWnd, CRR(0, 0, (picCross.Width), (picCross.Height), 14 * XMag, 14 * YMag), True

   Picture1.SetFocus
   
End Sub

Private Sub ShowSlider()
   If aLoaded Then
      
      SetStretchBltMode picSlider.hdc, HALFTONE
      
      picSlider.Cls
      StretchDIBits picSlider.hdc, _
      0, 0, (w1 * xmagslide), (h1 * ymagslide), _
      0, 0, w1, h1, _
      SliderMem(1, 1), bmslider, _
      DIB_RGB_COLORS, _
      vbSrcCopy
      
      SliderValue = 50
      Text1.Text = Str$(SliderValue)
      picSlider.Refresh
      
      With Picture1
         .Width = 2 + 0.8 * SliderValue * XMag
         .Height = 2 + 0.95 * SliderValue * YMag
         .Refresh
         SizeArr(iPicture1).xW = .Width
         SizeArr(iPicture1).yH = .Height
      End With
   
   End If
End Sub

Private Sub MaskSwitch2UP()
   SetStretchBltMode picSwitch2.hdc, HALFTONE
   picSwitch2.Cls
   
   ' Me background to picLSide
   BitBlt picSwitch2.hdc, 0, 0, picSwitch2.Width, picSwitch2.Height, _
   Me.hdc, picSwitch2.Left, picSwitch2.Top, vbSrcCopy

   ' Switch sprite
   StretchDIBits picSwitch2.hdc, _
   0, 0, picSwitch2.Width, picSwitch2.Height, _
   0, 0, w9, h9, _
   SwitchUpMaskMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picSwitch2.hdc, _
   0, 0, picSwitch2.Width, picSwitch2.Height, _
   0, 0, w9, h9, _
   SwitchUpMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   picSwitch2.Refresh

   imgLight.Visible = False
End Sub

Private Sub MaskSwitch2DN()
   SetStretchBltMode picSwitch2.hdc, HALFTONE
   picSwitch2.Cls
   
   ' Me background to picLSide
   BitBlt picSwitch2.hdc, 0, 0, picSwitch2.Width, picSwitch2.Height, _
   Me.hdc, picSwitch2.Left, picSwitch2.Top, vbSrcCopy

   ' Switch sprite
   StretchDIBits picSwitch2.hdc, _
   0, 0, picSwitch2.Width, picSwitch2.Height, _
   0, 0, w9, h9, _
   SwitchDnMaskMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picSwitch2.hdc, _
   0, 0, picSwitch2.Width, picSwitch2.Height, _
   0, 0, w9, h9, _
   SwitchDnMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   picSwitch2.Refresh

   imgLight.Visible = True
End Sub

Private Sub MaskSwitchUP()
   SetStretchBltMode picSwitch.hdc, HALFTONE
   picSwitch.Cls
   
   ' Me background to picLSide
   BitBlt picSwitch.hdc, 0, 0, picSwitch.Width, picSwitch.Height, _
   Me.hdc, picSwitch.Left, picSwitch.Top, vbSrcCopy

   ' Switch sprite
   StretchDIBits picSwitch.hdc, _
   0, 0, picSwitch.Width, picSwitch.Height, _
   0, 0, w9, h9, _
   SwitchUpMaskMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picSwitch.hdc, _
   0, 0, picSwitch.Width, picSwitch.Height, _
   0, 0, w9, h9, _
   SwitchUpMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   picSwitch.Refresh
End Sub

Private Sub MaskSwitchDN()
   SetStretchBltMode picSwitch.hdc, HALFTONE
   picSwitch.Cls
   
   ' Me background to picLSide
   BitBlt picSwitch.hdc, 0, 0, picSwitch.Width, picSwitch.Height, _
   Me.hdc, picSwitch.Left, picSwitch.Top, vbSrcCopy

   ' Switch sprite
   StretchDIBits picSwitch.hdc, _
   0, 0, picSwitch.Width, picSwitch.Height, _
   0, 0, w9, h9, _
   SwitchDnMaskMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picSwitch.hdc, _
   0, 0, picSwitch.Width, picSwitch.Height, _
   0, 0, w9, h9, _
   SwitchDnMem(1, 1), bmswitch, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   picSwitch.Refresh
End Sub

Private Sub MaskWatchIt()
   SetStretchBltMode picLSide.hdc, HALFTONE
   picLSide.Cls
   
   ' Me background to picLSide
   BitBlt picLSide.hdc, 0, 0, picLSide.Width, picLSide.Height, _
   Me.hdc, picLSide.Left, picLSide.Top, vbSrcCopy

   ' WatchIt sprite
   StretchDIBits picLSide.hdc, _
   0, 0, picLSide.Width, picLSide.Height, _
   0, 0, w4, h4, _
   WatchItMaskMem(1, 1), bmwatchit, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picLSide.hdc, _
   0, 0, picLSide.Width, picLSide.Height, _
   0, 0, w4, h4, _
   WatchItMem(1, 1), bmwatchit, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   picLSide.Refresh

End Sub

Private Sub MaskSlider()
Dim WD As Long
Dim WS As Long
   SetStretchBltMode picScale.hdc, HALFTONE
   
   ' Left side
   WD = (w2 * xmagslide - (picSlider.Left - picScale.Left))
   WS = (w2 - (picSlider.Left - picScale.Left) / xmagslide)
   
   StretchDIBits picScale.hdc, _
   picSlider.Left - picScale.Left, 0, WD, (h2 * ymagslide), _
   0, 0, WS, h2, _
   ScaleMem(1, 1), bmscale, _
   DIB_RGB_COLORS, _
   vbSrcCopy
   
   ' Right side
   WD = w2 * xmagslide - WD
   
   StretchDIBits picScale.hdc, _
   0, 0, WD, (h2 * ymagslide), _
   WS, 0, w2 - WS, h2, _
   ScaleMem(1, 1), bmscale, _
   DIB_RGB_COLORS, _
   vbSrcCopy
   
   picScale.Refresh
   
   ' picScale background to picSlider
   BitBlt picSlider.hdc, 0, 0, picSlider.Width, picSlider.Height, _
   picScale.hdc, picSlider.Left - picScale.Left, 0, vbSrcCopy

   ' Slider sprite
   StretchDIBits picSlider.hdc, _
   0, 0, picSlider.Width, picSlider.Height, _
   0, 0, w1, h1, _
   SliderMaskMem(1, 1), bmslider, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picSlider.hdc, _
   0, 0, picSlider.Width, picSlider.Height, _
   0, 0, w1, h1, _
   SliderMem(1, 1), bmslider, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   picSlider.Refresh
End Sub

Private Sub MaskVSlider()
Dim WL As Long
Dim WM As Long
Dim WR As Long

   
   ' Form background to picVScale
   BitBlt picVScale.hdc, 0, 0, picVScale.Width, picVScale.Height, _
   Me.hdc, picVScale.Left, picVScale.Top, vbSrcCopy
   
   picVScale.Refresh

   SetStretchBltMode picVScale.hdc, HALFTONE

   ' picVScale sprite
   StretchDIBits picVScale.hdc, _
   0, 0, picVScale.Width, picVScale.Height, _
   0, 0, w5, h5, _
   pvScaleMaskMem(1, 1), bmpvscale, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picVScale.hdc, _
   0, 0, picVScale.Width, picVScale.Height, _
   0, 0, w5, h5, _
   pvScaleMem(1, 1), bmpvscale, _
   DIB_RGB_COLORS, _
   vbSrcInvert
   
   picVScale.Refresh
   
   '-------------------
   
   WL = picVScale.Left - picVSlider.Left
   WM = picVScale.Width
   WR = picVSlider.Width - picVScale.Width - WL + 1

   ' Form & picVScale background to picVSlider
   
   BitBlt picVSlider.hdc, 0, 0, WL, picVSlider.Height, _
   Me.hdc, picVSlider.Left, picVSlider.Top, vbSrcCopy

   BitBlt picVSlider.hdc, WL, 0, WM, picVSlider.Height, _
   picVScale.hdc, 0, picVSlider.Top - picVScale.Top, vbSrcCopy

   BitBlt picVSlider.hdc, WL + WM, 0, WR, picVSlider.Height, _
   Me.hdc, picVScale.Left + picVScale.Width, picVSlider.Top, vbSrcCopy
   
   SetStretchBltMode picVSlider.hdc, HALFTONE
   
   ' Slider sprite
   StretchDIBits picVSlider.hdc, _
   0, 0, picVSlider.Width, picVSlider.Height, _
   0, 0, w6, h6, _
   pvSliderMaskMem(1, 1), bmpvslider, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picVSlider.hdc, _
   0, 0, picVSlider.Width, picVSlider.Height, _
   0, 0, w6, h6, _
   pvSliderMem(1, 1), bmpvslider, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   picVSlider.Refresh
End Sub


Private Sub MaskJustVSlider()
Dim WL As Long
Dim WM As Long
Dim WR As Long

   
   WL = picVScale.Left - picVSlider.Left
   WM = picVScale.Width
   WR = picVSlider.Width - picVScale.Width - WL

   ' Form & picVScale background to picVSlider
   'picVSlider.Cls
   
   BitBlt picVSlider.hdc, 0, 0, WL, picVSlider.Height, _
   Me.hdc, picVSlider.Left, picVSlider.Top, vbSrcCopy
   
   'picVSlider.Refresh

   BitBlt picVSlider.hdc, WL, 0, WM, picVSlider.Height, _
   picVScale.hdc, 0, picVSlider.Top - picVScale.Top, vbSrcCopy

   'picVSlider.Refresh

   BitBlt picVSlider.hdc, WL + WM, 0, WR, picVSlider.Height, _
   Me.hdc, picVScale.Left + picVScale.Width, picVSlider.Top, vbSrcCopy


   SetStretchBltMode picVSlider.hdc, HALFTONE

   ' Slider sprite
   StretchDIBits picVSlider.hdc, _
   0, 0, picVSlider.Width, picVSlider.Height, _
   0, 0, w6, h6, _
   pvSliderMaskMem(1, 1), bmpvslider, _
   DIB_RGB_COLORS, _
   vbSrcAnd

   StretchDIBits picVSlider.hdc, _
   0, 0, picVSlider.Width, picVSlider.Height, _
   0, 0, w6, h6, _
   pvSliderMem(1, 1), bmpvslider, _
   DIB_RGB_COLORS, _
   vbSrcInvert

   
   picVSlider.Refresh
End Sub




Private Sub ShowWholePicture()
'Public W As Long, H As Long
'Public XMag,YMag

' Necessary to do it twice sometimes?

   If aLoaded Then
   
      SetStretchBltMode Me.hdc, HALFTONE
      
      StretchDIBits Me.hdc, _
      0, 0, (W * XMag), (H * YMag), _
      0, 0, W, H, _
      ArrMem(1, 1), bm, _
      DIB_RGB_COLORS, _
      vbSrcCopy
   
      Me.Refresh
      
      StretchDIBits Me.hdc, _
      0, 0, (W * XMag), (H * YMag), _
      0, 0, W, H, _
      ArrMem(1, 1), bm, _
      DIB_RGB_COLORS, _
      vbSrcCopy
   
      Me.Refresh
      
      DoEvents
   
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Form_Unload 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Make sure all forms cleared
Dim Form As Form
   Erase ArrMem
   sndPlaySound "", SND_PURGE
   
   ' Collapse form
   Do
      Me.Height = Me.Height - 30
      Me.Top = (Screen.Height - Me.Height) \ 2
      DoEvents
      If Me.Height <= 100 Then Exit Do
   Loop
   Do
      Me.Width = Me.Width - 30
      Me.Left = (Screen.Width - Me.Width) \ 2
      DoEvents
      If Me.Width <= 500 Then Exit Do
   Loop
   
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wParam As Long
   
   With Me
      If X > .Width \ STX - 24 * XMag And X < .Width \ STX Then
      If Y > .Height \ STY - 24 * YMag And Y < .Height \ STY Then
         wParam = HTBOTTOMRIGHT
         If wParam Then
            
            Call ReleaseCapture
            Call SendMessage(.hWnd, WM_NCLBUTTONDOWN, wParam, 0)
      
            If .Width < 256 * STX Then .Width = 256 * STX
            If .Height < 256 * STY Then .Height = 256 * STY
      
            XMag = .Width / ORGFW
            YMag = .Height / ORGFH
            xmagslide = (.Width / STX) / FWO
            ymagslide = (.Height / STY) / FHO
            
            If aChecked Then
               SWR hWnd, CRR(6, 6, (FWO * XMag), (FHO * YMag), 50, 50), True
            Else
               SWR hWnd, CRR(0, 0, (FWO * XMag), (FHO * YMag), 0, 0), True
            End If
            
            ResizePicture1
            
            If Not aTile Then ShowWholePicture
            ResizeControls
            MaskSlider
            MaskWatchIt
            MaskVSlider
            MaskSwitchUP
            MaskSwitch2UP
            
            SWidth = GetSystemMetrics(SM_CXSCREEN) * STX
            SHeight = GetSystemMetrics(SM_CYSCREEN) * STY
            'Centre form
            .Top = (SHeight - .Height) / 2
            .Left = (SWidth - .Width) / 2
            
            VSliderValue = 31
            LabVScale = LTrim$(Str$(VSliderValue))
            
            StretchDIBits picWheel.hdc, _
            0, 0, picWheel.Width, picWheel.Height, _
            VSliderValue * 19, 0, 19, 19, _
            WheelMem(1, 1), bmwheel, _
            DIB_RGB_COLORS, _
            vbSrcCopy
            picWheel.Refresh
            
            picCross_MouseUp 1, 0, 0, 0
            
         End If
      End If
      End If
   End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   With Me
       If X > .Width \ STX - 24 * XMag And X < .Width \ STX Then
         
         If Y > .Height \ STY - 24 * YMag And Y < .Height \ STY Then
             .MousePointer = vbSizeNWSE
         Else
             .MousePointer = vbDefault
         End If
       
       Else
          .MousePointer = vbDefault
       End If
   End With
End Sub



