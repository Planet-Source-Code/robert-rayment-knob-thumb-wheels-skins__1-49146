Attribute VB_Name = "Module1"
Option Explicit

Public WStart As Long   ' Starting image width
Public HStart As Long   ' & height
Public MAG As Long      ' Mag onto picStrip
Public NRI As Long      ' Number of repeated images
Public zRAD As Single   ' Radius of rotation about centre
Public zANG As Single   ' Angle decrement to rotate each image in strip
Public aAntiAlias As Boolean

Public WStrip           ' Strip width
Public HStrip           ' Strip height
Public xcen As Single   ' = WStart/2
Public ycen As Single   ' = HStart/2
Public xSrc As Single   ' Source & destination coords
Public ySrc As Single
Public xdes As Single
Public ydes As Single

Public bStartMem() As Byte    ' BGRA @ X,Y
Public bRotatedMem() As Byte  ' BGRA @ X,Y rotated from bStartMem()
Public bStripMem() As Byte    ' BGRA @ X,Y
Public bTestMem() As Long     ' @ X,Y
Public zFrame As Single       ' For testing wheel
Public aBuildHorizontal As Boolean
Public aknobwheel  As Boolean ' knob or thumb

Public aLoaded As Boolean
' General Public
Public i As Long
Public j As Long
Public k As Long
Public a$
Public Const pi# = 3.14159297
Public Const DTR# = pi# / 180 ' deg to radians


Public Sub Rotate(N As Long)
'Public xcen As Single   ' = WStart/2
'Public ycen As Single   ' = HStart/2
'Public xsrc As Single   ' Source & destination coords
'Public ysrc As Single
'Public xdes As Single
'Public ydes As Single
'Public zRAD As Long      ' Radius of rotation about centre
'Public zANG As Single   ' Angle decrement to rotate each image in strip
Dim zTotAngle As Single
Dim zcos As Single
Dim zsin As Single
Dim ixs As Long
Dim iys As Long
Dim jlo As Long, jhi As Long
Dim ilo As Long, ihi As Long
Dim zRAD2 As Single

   ' bStartMem() rotated zANG*(N-1) to bRotatedMem()
   CopyMemory bRotatedMem(1, 1, 1), bStartMem(1, 1, 1), 4 * WStart * HStart
   
   If zANG = 0 Or zRAD = 0 Then
   
   Else
      
      zRAD2 = zRAD * zRAD
      zTotAngle = DTR# * zANG * (N - 1)
      zcos = Cos(zTotAngle)
      zsin = Sin(-zTotAngle)
      
      xcen = (WStart + 1) / 2
      ycen = (HStart + 1) / 2

      jlo = ycen - zRAD: If jlo < 1 Then jlo = 1
      jhi = ycen + zRAD: If jhi > HStart Then jhi = HStart
      ilo = xcen - zRAD: If ilo < 1 Then ilo = 1
      ihi = xcen + zRAD: If ihi > WStart Then ihi = WStart
      
      For j = jlo To jhi
      For i = ilo To ihi
         ' Find source point from rotated destination point
         ixs = xcen + (i - xcen) * zcos - (j - ycen) * zsin
         iys = ycen + (j - ycen) * zcos + (i - xcen) * zsin
      
         If ixs > 0 Then
         If ixs <= WStart Then
         If iys > 0 Then
         If iys <= HStart Then
         If (iys - ycen) * (iys - ycen) + (ixs - xcen) * (ixs - xcen) <= zRAD2 Then
            bRotatedMem(1, i, j) = bStartMem(1, ixs, iys)
            bRotatedMem(2, i, j) = bStartMem(2, ixs, iys)
            bRotatedMem(3, i, j) = bStartMem(3, ixs, iys)
         End If
         End If
         End If
         End If
         End If
      Next i
      Next j
   
   End If
End Sub

Public Sub AARotate(N As Long)
Dim zTotAngle As Single
Dim zcos As Single
Dim zsin As Single
Dim ixs As Long
Dim iys As Long
Dim jlo As Long, jhi As Long
Dim ilo As Long, ihi As Long
Dim zRAD2 As Single
Dim xs As Single
Dim ys As Single
Dim xsf As Single
Dim ysf As Single
Dim zsf As Single
Dim culB As Long
Dim culG As Long
Dim culR As Long
Dim culB1 As Long
Dim culG1 As Long
Dim culR1 As Long

   ' bStartMem() anti-alias rotated zANG*(N-1)to bRotatedMem()
   CopyMemory bRotatedMem(1, 1, 1), bStartMem(1, 1, 1), 4 * WStart * HStart
   
   If zANG = 0 Or zRAD = 0 Then
      'AA bRotatedMem()
   
   Else
      zTotAngle = zANG * (N - 1)
      zcos = Cos(zTotAngle)
      zsin = Sin(-zTotAngle)
      
      zRAD2 = zRAD * zRAD
      zTotAngle = DTR# * zANG * (N - 1)
      zcos = Cos(zTotAngle)
      zsin = Sin(-zTotAngle)
      
      xcen = (WStart + 1) / 2
      ycen = (HStart + 1) / 2

      jlo = ycen - zRAD: If jlo < 1 Then jlo = 1
      jhi = ycen + zRAD: If jhi > HStart Then jhi = HStart
      ilo = xcen - zRAD: If ilo < 1 Then ilo = 1
      ihi = xcen + zRAD: If ihi > WStart Then ihi = WStart
      
      For j = jlo To jhi
      For i = ilo To ihi
      
         ' Find source point from rotated destination point
         xs = xcen + (i - xcen) * zcos - (j - ycen) * zsin
         ys = ycen + (j - ycen) * zcos + (i - xcen) * zsin
      
         ' Bottom left coords of bounding rectangle
         ixs = Int(xs)
         iys = Int(ys)

         If ixs > 1 Then
         If ixs < WStart Then
         If iys > 1 Then
         If iys < HStart Then
         If (iys - ycen) * (iys - ycen) + (ixs - xcen) * (ixs - xcen) <= zRAD2 Then
            ' Scale factors
            xsf = xs - ixs
            ysf = ys - iys
            zsf = (1 - xsf)
            
            ' Weight along bottom x axis
            ' bStartMem(1, ixs, iys) bStartMem(1, ixs+1, iys)
            culB = zsf * bStartMem(1, ixs, iys) + xsf * bStartMem(1, ixs + 1, iys)
            culG = zsf * bStartMem(2, ixs, iys) + xsf * bStartMem(2, ixs + 1, iys)
            culR = zsf * bStartMem(3, ixs, iys) + xsf * bStartMem(3, ixs + 1, iys)
      
            ' Weight along top x axis
            ' bStartMem(1, ixs, iys+1) bStartMem(1, ixs+1, iys+1)
            culB1 = zsf * bStartMem(1, ixs, iys + 1) + xsf * bStartMem(1, ixs + 1, iys + 1)
            culG1 = zsf * bStartMem(2, ixs, iys + 1) + xsf * bStartMem(2, ixs + 1, iys + 1)
            culR1 = zsf * bStartMem(3, ixs, iys + 1) + xsf * bStartMem(3, ixs + 1, iys + 1)
      
            ' Weight along y axis
            culB = (1 - ysf) * culB + ysf * culB1
            culG = (1 - ysf) * culG + ysf * culG1
            culR = (1 - ysf) * culR + ysf * culR1
      
            If culB > 255 Then culB = 255
            If culG > 255 Then culG = 255
            If culR > 255 Then culR = 255
      
            If culB < 0 Then culB = 0
            If culG < 0 Then culG = 0
            If culR < 0 Then culR = 0
      
            bRotatedMem(1, i, j) = culB
            bRotatedMem(2, i, j) = culG
            bRotatedMem(3, i, j) = culR
         End If
         End If
         End If
         End If
         End If
      Next i
      Next j
   
   End If
End Sub


Public Sub FixScrollbars(picC As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)
   ' picC = Container
   ' picP = Picture
      
      HS.Max = picP.Width - picC.Width + 12   ' +4 to allow for border
      HS.LargeChange = picC.Width \ 10
      HS.SmallChange = 1
      
      HS.Top = picC.Top + picC.Height + 1
      
      HS.Left = picC.Left
      HS.Width = picC.Width
      If picP.Width < picC.Width Then
         HS.Visible = False
         'HS.Enabled = False
      Else
         HS.Visible = True
         'HS.Enabled = True
      End If
      
      VS.Max = picP.Height - picC.Height + 12 ' +4 to allow for border
      VS.LargeChange = picC.Height \ 10
      VS.SmallChange = 1
      VS.Top = picC.Top
      
      VS.Left = picC.Left - VS.Width - 1
      
      VS.Height = picC.Height
      If picP.Height < picC.Height Then
         VS.Visible = False
         'VS.Enabled = False
      Else
         VS.Visible = True
         'VS.Enabled = True
      End If
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


