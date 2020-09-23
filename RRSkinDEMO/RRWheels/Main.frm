VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   " Thumb & Knob Wheel Creation  by Robert Rayment"
   ClientHeight    =   7800
   ClientLeft      =   165
   ClientTop       =   30
   ClientWidth     =   11880
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   7185
      MousePointer    =   2  'Cross
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   26
      Top             =   2610
      Width           =   1590
   End
   Begin VB.HScrollBar HSTest 
      Height          =   210
      Left            =   7275
      TabIndex        =   25
      Top             =   6465
      Width           =   3225
   End
   Begin VB.Frame fraMag 
      Caption         =   "Magnification"
      Height          =   525
      Left            =   7260
      TabIndex        =   21
      Top             =   6750
      Width           =   3915
      Begin VB.CommandButton cmdTEST 
         Caption         =   "Test thumb"
         Height          =   225
         Index           =   1
         Left            =   2565
         TabIndex        =   27
         Top             =   210
         Width           =   1050
      End
      Begin VB.CommandButton cmdTEST 
         Caption         =   "Test knob"
         Height          =   225
         Index           =   0
         Left            =   1125
         TabIndex        =   23
         Top             =   210
         Width           =   1050
      End
      Begin VB.HScrollBar HSMag 
         Height          =   195
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   22
         Top             =   240
         Value           =   1
         Width           =   465
      End
      Begin VB.Label LabMag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   690
         TabIndex        =   24
         Top             =   225
         Width           =   240
      End
   End
   Begin VB.VScrollBar VS_Strip 
      Height          =   4215
      Left            =   105
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2370
      Width           =   180
   End
   Begin VB.Frame fraAction 
      Height          =   1980
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   5205
      Begin VB.CheckBox chkAA 
         Caption         =   "Anti-alias"
         Height          =   240
         Left            =   3585
         TabIndex        =   18
         Top             =   1275
         Width           =   945
      End
      Begin VB.CommandButton cmdStrip 
         Caption         =   "Build vertical strip"
         Height          =   240
         Index           =   1
         Left            =   3345
         TabIndex        =   13
         Top             =   885
         Width           =   1620
      End
      Begin VB.CommandButton cmdStrip 
         Caption         =   "Build horizontal strip"
         Height          =   240
         Index           =   0
         Left            =   3345
         TabIndex        =   12
         Top             =   525
         Width           =   1620
      End
      Begin VB.TextBox txtIN 
         Height          =   285
         Index           =   2
         Left            =   2625
         MaxLength       =   5
         TabIndex        =   11
         Text            =   "txtIN"
         Top             =   1215
         Width           =   555
      End
      Begin VB.TextBox txtIN 
         Height          =   285
         Index           =   1
         Left            =   2625
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "txtIN"
         Top             =   870
         Width           =   540
      End
      Begin VB.TextBox txtIN 
         Height          =   285
         Index           =   0
         Left            =   2610
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "txtIN"
         Top             =   540
         Width           =   555
      End
      Begin VB.Label LabXY 
         Caption         =   "Strip W =  H ="
         Height          =   210
         Index           =   2
         Left            =   3270
         TabIndex        =   19
         Top             =   1680
         Width           =   2565
      End
      Begin VB.Label LabXY 
         Caption         =   "LabXY(1)"
         Height          =   240
         Index           =   1
         Left            =   1680
         TabIndex        =   17
         Top             =   1665
         Width           =   1350
      End
      Begin VB.Label LabXY 
         Caption         =   "LabXY(0)"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   1650
         Width           =   1320
      End
      Begin VB.Label LabMove 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Move frame"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   165
         MousePointer    =   5  'Size
         TabIndex        =   14
         Top             =   195
         Width           =   4785
      End
      Begin VB.Label Label1 
         Caption         =   "Rotation angle - zANG deg              (eg 360/NRI)"
         Height          =   405
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1215
         Width           =   2025
      End
      Begin VB.Label Label1 
         Caption         =   "Radius of rotation - zRAD pixels"
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   8
         Top             =   915
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "Num repeated images - NRI"
         Height          =   270
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   600
         Width           =   2100
      End
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   315
      ScaleHeight     =   334
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   2370
      Width           =   6300
      Begin VB.PictureBox picStrip 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   15
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   52
         TabIndex        =   1
         Top             =   15
         Width           =   780
      End
   End
   Begin VB.HScrollBar HS_Strip 
      Height          =   180
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7425
      Width           =   5865
   End
   Begin VB.PictureBox picStart 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1200
      Left            =   315
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   2
      Top             =   420
      Width           =   1200
   End
   Begin VB.Label LabValue 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabValue"
      Height          =   225
      Left            =   10575
      TabIndex        =   29
      Top             =   6480
      Width           =   630
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Image strip to save"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2880
      TabIndex        =   28
      Top             =   2115
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      BorderWidth     =   12
      Height          =   5130
      Left            =   6900
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   4605
   End
   Begin VB.Label LabWHStart 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W =  H =  File: "
      Height          =   255
      Left            =   315
      TabIndex        =   4
      Top             =   75
      Width           =   1125
   End
   Begin VB.Label LabInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NB This program requires a starting image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5760
      TabIndex        =   20
      Top             =   360
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuFiles 
         Caption         =   "&Load start image"
         Index           =   0
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "&Save image strip"
         Index           =   2
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "&Exit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Thumb & Knob Wheel Creation by Robert Rayment Oct 2003

Dim PathSpec$, CurrentPath$
Dim FileSpec$

Dim xfra As Single
Dim yfra As Single
Dim STX As Long
Dim STY As Long

Dim zprevAngle As Single
Dim FrameValue As Long

Private CommonDialog1 As New OSDialog

Private Sub Form_Load()

   a$ = vbCr
   a$ = a$ + "    NB This program needs a starting image.  A knob wheel should  " & vbCr
   a$ = a$ + "    be square with odd numbered width && height to a maximum of " & vbCr
   a$ = a$ + "    121 x 121 pixels.  Try TestWheel.bmp first.  The wheel can be " & vbCr
   a$ = a$ + "    rotated from the scroll bar or by using the mouse over the wheel. " & vbCr
   LabInfo = a$
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrentPath$ = PathSpec$
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   ' Initial values
   MAG = 1
   NRI = 36 '1
   zRAD = 16 '0
   zANG = 10 '0
   
   HSMag.Value = MAG
   txtIN(0) = NRI
   txtIN(1) = zRAD
   txtIN(2) = zANG
   
   LabXY(0) = "0, 0"
   LabXY(1) = "0, 0"
   
   picStrip.Left = 0
   picStrip.Top = 0
   
   FixScrollbars picFrame, picStrip, HS_Strip, VS_Strip
   fraMag.Visible = False
   fraAction.Visible = False
   aLoaded = False
   aAntiAlias = False
   
   ReDim bStartMem(1 To 4, 1 To 1, 1 To 1)
   ReDim bRotatedMem(1 To 4, 1 To 1, 1 To 1)
   ReDim bStripMem(1 To 4, 1 To 1, 1 To 1)
   ReDim bTestMem(1 To 1, 1 To 1)
End Sub

'#### STRIP ACTION ################################################
Private Sub chkAA_Click()
   aAntiAlias = Not aAntiAlias
End Sub

'### Rotate wheel with mouse ###############################################

Private Sub picTest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xc As Single
Dim yc As Single
Dim zAngle As Single

'Dim zprevAngle As Single
'Dim FrameValue As Long

   If aknobwheel Then
      If Button = vbLeftButton Then
         xc = picTest.Width / 2
         yc = picTest.Height / 2
         zAngle = zAtn2(Y - yc, X - xc)
         If zAngle > zprevAngle Then   ' Clockwise +
            If FrameValue < NRI Then FrameValue = FrameValue + 1
         Else     ' Anti-clockwise
            If FrameValue > 0 Then FrameValue = FrameValue - 1
         End If
   
         HSTest.Max = NRI
         HSTest.Min = 1
         If FrameValue < 1 Then
            FrameValue = 1
         ElseIf FrameValue > NRI Then
            FrameValue = NRI
         End If
         
         HSTest.Value = FrameValue  ' Goes to Sub HSTest_Change
                                    'NB Could use BitBlt instead.
         'EG  without magnification!
         'BitBlt picTest.hdc, 0, 0, WStart, HStart, _
         'picStrip.hdc, FrameValue * WStart, 0, vbSrcCopy
         
         zprevAngle = zAngle
      End If
   End If
End Sub

Private Sub txtIN_Change(Index As Integer)
   Select Case Index
   Case 0   ' NRI
      NRI = Val(txtIN(0))
   Case 1   ' zRAD
      zRAD = Val(txtIN(1))
   Case 2   ' zANG
      zANG = Val(txtIN(2))
   End Select
End Sub

Private Sub cmdStrip_Click(Index As Integer)
'Public WStart As Long   ' Starting image width
'Public HStart As Long   ' & height
'Public NRI As Long      ' Number of repeated images
'Public WStrip           ' Strip width
'Public HStrip           ' Strip height

   If NRI < 1 Then txtIN(0).Text = "1"
   If zRAD < 0 Then txtIN(1).Text = "1"
   If zANG = 0 Then txtIN(2).Text = "0"
   
   Select Case Index
   Case 0   ' Build horizontal strip
      WStrip = NRI * WStart
      HStrip = HStart
      
      picStrip.Picture = LoadPicture
      picStrip.Refresh
      
      With picStrip
         .Width = WStrip
         .Height = HStrip
      End With
      picStrip.Refresh
      DoEvents
      
      FixScrollbars picFrame, picStrip, HS_Strip, VS_Strip
      LabXY(2) = "Strip:  W =" & Str$(WStrip) & "  H =" & Str$(HStrip) & " "
      
      BUILD_HORIZONTAL_STRIP
      
      aBuildHorizontal = True
   
   Case 1   ' Build vertical strip
      WStrip = WStart
      
      If NRI * HStart > 8000 Then
         NRI = 8000 \ HStart
         MsgBox " Too large for vertical strip. " & vbCr _
         & " Reduce NRI so that NRI x Start width <= 8000" & vbCr _
         & " NRI reset to" & Str$(NRI), vbInformation, " "
         txtIN(0) = NRI
         txtIN(0).Refresh
      End If
      
      HStrip = NRI * HStart
      
      picStrip.Picture = LoadPicture
      picStrip.Refresh
      
      With picStrip
         .Width = WStrip
         .Height = HStrip
      End With
      picStrip.Refresh
      DoEvents
      
      FixScrollbars picFrame, picStrip, HS_Strip, VS_Strip
      LabXY(2) = "Strip:  W =" & Str$(WStrip) & "  H =" & Str$(HStrip) & " "
      
      BUILD_VERTICAL_STRIP
      
      aBuildHorizontal = False
   
   End Select
   
   With picTest
      .Width = WStart
      .Height = HStart
      .Picture = LoadPicture
   End With
   
   fraMag.Visible = True
End Sub

Private Sub BUILD_HORIZONTAL_STRIP()
Dim N As Long
'bStartMem(1 To 4, 1 To WStart, 1 To HStart)
   
   ReDim bRotatedMem(1 To 4, 1 To WStart, 1 To HStart)
   
   For N = 1 To NRI
      If aAntiAlias Then
         AARotate N  ' bStartMem() to bRotatedMem()
      Else
         Rotate N    ' bStartMem() to bRotatedMem()
      End If
      
      SetStretchBltMode picStrip.hdc, HALFTONE

      ' Locate rotated image on picStrip
      StretchDIBits picStrip.hdc, _
      (N - 1) * WStart, 0, WStart, HStart, _
      0, 0, WStart, HStart, _
      bRotatedMem(1, 1, 1), bmStart, _
      DIB_RGB_COLORS, vbSrcCopy
   Next N
   picStrip.Refresh
End Sub

Private Sub BUILD_VERTICAL_STRIP()
Dim N As Long
'bStartMem(1 To 4, 1 To WStart, 1 To HStart)
   
   ReDim bRotatedMem(1 To 4, 1 To WStart, 1 To HStart)
   
   For N = 1 To NRI
      If aAntiAlias Then
         AARotate N  ' bStartMem() to bRotatedMem()
      Else
         Rotate N    ' bStartMem() to bRotatedMem()
      End If
      
      SetStretchBltMode picStrip.hdc, HALFTONE
      
      ' Locate rotated image on picStrip
      StretchDIBits picStrip.hdc, _
      0, (N - 1) * HStart, WStart, HStart, _
      0, 0, WStart, HStart, _
      bRotatedMem(1, 1, 1), bmStart, _
      DIB_RGB_COLORS, vbSrcCopy
   Next N
   picStrip.Refresh
End Sub
'#### END STRIP ACTION ################################################

'#### TEST ###########################################
Private Sub cmdTEST_Click(Index As Integer)
Dim TW As Long
Dim TH As Long
   HSTest.Enabled = False
   
   ' Prevent picTest over-sizing
   Do
      TW = WStart * MAG
      TH = HStart * MAG
      If TW <= 256 And TH <= 256 Then Exit Do
      MAG = MAG - 1
   Loop
   HSMag.Value = MAG
   
   SetStretchBltMode picTest.hdc, HALFTONE
   
   Select Case Index
   Case 0   ' Test knob wheel
      aknobwheel = True
      
      With picTest
         .Width = WStart * MAG
         .Height = HStart * MAG
         .Picture = LoadPicture
      End With
      HSTest.Max = NRI
      HSTest.Min = 1
      HSTest.LargeChange = 1
      HSTest.Value = 1
      zFrame = 1
      
      ' Get strip image to bStripMem for testing
      ReDim bStripMem(1 To 4, 1 To WStrip, 1 To HStrip)
      FillBMPStruc bStripMem(), bmStrip
      GETDIB picStrip.Image, bStripMem(), bmStrip
      
      If aBuildHorizontal Then
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, 0, WStart, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      Else
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, (NRI - 1) * HStart, WStart, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      End If
      picTest.Refresh
   
   Case 1   ' Test thumb wheel
      aknobwheel = False
   
      If aBuildHorizontal Then
         With picTest
            .Width = WStart * MAG
            .Height = HStart
            .Picture = LoadPicture
         End With
      Else
         With picTest
            .Width = WStart
            .Height = HStart * MAG
            .Picture = LoadPicture
         End With
      End If
      HSTest.Max = (NRI - MAG) * 4
      HSTest.Min = 1
      HSTest.LargeChange = 1
      HSTest.Value = 1
      zFrame = 0.25
      
      ' Get strip image to bStripMem for testing
      ReDim bStripMem(1 To 4, 1 To WStrip, 1 To HStrip)
      FillBMPStruc bStripMem(), bmStrip
      GETDIB picStrip.Image, bStripMem(), bmStrip
      
      If aBuildHorizontal Then
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, 0, WStart * MAG, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      Else
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, (NRI - zFrame - MAG) * HStart, WStart, HStart * MAG, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      End If
      picTest.Refresh
   
   End Select
   HSTest.Enabled = True
End Sub

Private Sub HSMag_Change()
   HSTest.Enabled = False
   MAG = HSMag.Value
   LabMag = Trim$(Str$(MAG))
End Sub

Private Sub HSTest_Change()

   SetStretchBltMode picTest.hdc, HALFTONE

   If aknobwheel Then
      zFrame = HSTest.Value
      If aBuildHorizontal Then
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         (zFrame - 1) * WStart, 0, WStart, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      Else
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, (NRI - zFrame) * HStart, WStart, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      End If
   Else  ' Thumb wheel
      zFrame = HSTest.Value / 4
      If aBuildHorizontal Then
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         (zFrame - 0.25) * WStart, 0, WStart * MAG, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      Else
         picTest.Cls
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, (NRI - zFrame - MAG) * HStart, WStart, HStart * MAG, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      End If
   End If
   picTest.Refresh
   LabValue = Str$(zFrame)

End Sub

Private Sub HSTest_Scroll()
   
   SetStretchBltMode picTest.hdc, HALFTONE
   
   If aknobwheel Then
      zFrame = HSTest.Value
      If aBuildHorizontal Then
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         (zFrame - 1) * WStart, 0, WStart, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      Else
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, (NRI - zFrame) * HStart, WStart, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      End If
   Else  ' Thumb wheel
      zFrame = HSTest.Value / 4
      If aBuildHorizontal Then
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         (zFrame - 0.25) * WStart, 0, WStart * MAG, HStart, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      Else
         picTest.Cls
         StretchDIBits picTest.hdc, _
         0, 0, picTest.Width, picTest.Height, _
         0, (NRI - zFrame - MAG) * HStart, WStart, HStart * MAG, _
         bStripMem(1, 1, 1), bmStrip, _
         DIB_RGB_COLORS, _
         vbSrcCopy
      End If
   End If
   
   picTest.Refresh
   LabValue = Str$(zFrame)
End Sub
'#### END TEST ###########################################

'#### PICTURE SCROLL BAR ##############################
Private Sub HS_Strip_Change()
   picStrip.Left = -HS_Strip.Value
End Sub

Private Sub HS_Strip_Scroll()
   picStrip.Left = -HS_Strip.Value
End Sub

Private Sub VS_Strip_Change()
   picStrip.Top = -VS_Strip.Value
End Sub

Private Sub VS_Strip_Scroll()
   picStrip.Top = -VS_Strip.Value
End Sub
'#### END PICTURE SCROLL BAR ##############################


Private Sub LabMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Move fraAction
   xfra = X
   yfra = Y
End Sub

Private Sub LabMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Move fraAction
   If Button = vbLeftButton Then
      fraAction.Left = fraAction.Left + (X - xfra) / STX
      fraAction.Top = fraAction.Top + (Y - yfra) / STY
   End If
End Sub

Private Sub LabMove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.Refresh
End Sub

Private Sub picStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabXY(0) = "X, Y =" & Str$(X) & " ," & Str$(Y)
End Sub

Private Sub picStrip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabXY(1) = "X, Y =" & Str$(X) & " ," & Str$(Y)
End Sub

'#### LOAD & SAVE #####################################################
Private Sub mnuFiles_Click(Index As Integer)
Dim Title$, Filt$, InDir$
   Set CommonDialog1 = New OSDialog
   
   Select Case Index
   Case 0
   '   LOAD
      Title$ = "Load a picture file"
      Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
      InDir$ = CurrentPath$ 'Pathspec$
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
      If Len(FileSpec$) <> 0 Then
         CurrentPath$ = FileSpec$
         picStart.Picture = LoadPicture(FileSpec$)
         WStart = picStart.Width - 4   ' -4 for border
         HStart = picStart.Height - 4
         
         If WStart > 121 Or HStart > 121 Then
            MsgBox " Maximum height & width 121 x 121 " & vbCr & _
            " This W x H is" & Str$(WStart) & " x" & Str$(HStart), vbInformation, " "
            WStart = 121
            HStart = 121
            With picStart
               .Picture = LoadPicture
               .Width = WStart + 4
               .Height = HStart + 4
               .Refresh
            End With
            Exit Sub
         End If
         
         fraAction.Visible = True
         aLoaded = True
         
         LabWHStart = " W =" & Str$(WStart) & "  H =" & Str$(HStart) & "  File: " & FileSpec$ & " "
         
         ReDim bStartMem(1 To 4, 1 To WStart, 1 To HStart)
         FillBMPStruc bStartMem(), bmStart
         GETDIB picStart.Image, bStartMem(), bmStart
         ReDim bRotatedMem(1 To 4, 1 To WStart, 1 To HStart)

         MAG = 1
         HSMag.Value = MAG
         WStrip = WStart
         HStrip = HStart
         picStrip.Width = WStrip
         picStrip.Height = HStrip
         LabXY(2) = "Strip:  W =" & Str$(WStrip) & "  H =" & Str$(HStrip) & " "
         fraMag.Visible = False
      End If
   Case 1   ' --
   
   Case 2
   '   SAVE
      If aLoaded = True Then
         Title$ = "Save As bmp"
         Filt$ = "Save bmp|*.bmp"
         InDir$ = CurrentPath$ 'Pathspec$
         FixExtension FileSpec$, ".bmp"
         CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
      
         If Len(FileSpec$) <> 0 Then
            CurrentPath$ = FileSpec$
            FixExtension FileSpec$, ".bmp"
            SavePicture picStrip.Image, FileSpec$
         End If
      End If
   Case 3   ' --
   Case 4   ' Exit
      Form_Unload 0
   End Select
   Set CommonDialog1 = Nothing
End Sub
'#### END LOAD & SAVE #####################################################

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

