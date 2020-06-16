Attribute VB_Name = "VB3DDraw"
Public Type CRGB
 R As Integer
 G As Integer
 B As Integer
End Type

Public Type P2D
 X As Integer
 Y As Integer
 Z As Double
 Color As CRGB
End Type

Public Type P3D
 X As Double
 Y As Double
 Z As Double
 Color As CRGB
End Type

Public Type Triangle
 Point2D(0 To 2) As P2D
 Point3D(0 To 2) As P3D
End Type

Public Camera As Double, FOV As Double, RenderDistance As Double, ZBuffer() As Double, AntiMix As Double

Const PI = 3.1415926

Public Sub Draw_Init(p As PictureBox)
 p.AutoRedraw = True
 p.BackColor = RGB(0, 0, 0)
 p.ForeColor = RGB(255, 255, 255)
 p.ScaleMode = 3
 ReDim ZBuffer(p.ScaleWidth * p.ScaleHeight)
End Sub

Public Sub Draw_Clear(p As PictureBox)
 p.Cls
 cr = Camera - RenderDistance
 For i = 0 To UBound(ZBuffer)
  ZBuffer(i) = cr
 Next i
End Sub

Public Sub MPST(p As PictureBox, t As Triangle, Optional offX As Double = 0, Optional offY As Double = 0, Optional SizeUp As Double = 1)
 s = Sqr(p.ScaleWidth ^ 2 + p.ScaleHeight ^ 2) / 2 / Tan(FOV * PI / 360)
 For i = 0 To 2
  k = s / (s + Camera - t.Point3D(i).Z)
  t.Point2D(i).X = Fix(p.ScaleWidth / 2 + k * (t.Point3D(i).X * SizeUp + offX))
  t.Point2D(i).Y = Fix(p.ScaleHeight / 2 - k * (t.Point3D(i).Y * SizeUp + offY))
  t.Point2D(i).Z = t.Point3D(i).Z * SizeUp
  t.Point2D(i).Color = t.Point3D(i).Color
 Next i
End Sub

Public Sub Rotate2D(X As Double, Y As Double, Angle As Double)
 tX = X
 tY = Y
 X = tX * Cos(Angle * PI / 180) - tY * Sin(Angle * PI / 180)
 Y = tX * Sin(Angle * PI / 180) + tY * Cos(Angle * PI / 180)
End Sub

Public Sub DrawSimpleTriangle(p As PictureBox, p1 As P2D, p2 As P2D, p3 As P2D)
 Dim x1 As Integer, x2 As Integer, dy As Integer
 Dim k1 As Double, k2 As Double
 Dim kz1 As Double, kz2 As Double, kzX As Double
 Dim kr1 As Double, kg1 As Double, kb1 As Double
 Dim kr2 As Double, kg2 As Double, kb2 As Double
 Dim krX As Double, kgX As Double, kbX As Double
 Dim tr As Integer, tg As Integer, tb As Integer, tz As Double
 Dim c1 As CRGB, c2 As CRGB
 Dim tmp As P2D
 If p1.Y = p2.Y And p2.Y = p3.Y Then
  Exit Sub
 End If
 If p1.Y = p2.Y Then
  k1 = (p1.X - p3.X) / (p1.Y - p3.Y)
  k2 = (p2.X - p3.X) / (p2.Y - p3.Y)
  kz1 = (p1.Z - p3.Z) / (p1.Y - p3.Y)
  kz2 = (p2.Z - p3.Z) / (p2.Y - p3.Y)
  kr1 = (p1.Color.R - p3.Color.R) / (p1.Y - p3.Y)
  kg1 = (p1.Color.G - p3.Color.G) / (p1.Y - p3.Y)
  kb1 = (p1.Color.B - p3.Color.B) / (p1.Y - p3.Y)
  kr2 = (p2.Color.R - p3.Color.R) / (p2.Y - p3.Y)
  kg2 = (p2.Color.G - p3.Color.G) / (p2.Y - p3.Y)
  kb2 = (p2.Color.B - p3.Color.B) / (p2.Y - p3.Y)
  For i = p1.Y To p3.Y
   dy = i - p1.Y
   x1 = Fix(p1.X + dy * k1)
   x2 = Fix(p2.X + dy * k2)
   z1 = p1.Z + dy * kz1
   z2 = p2.Z + dy * kz2
   c1.R = Fix(p1.Color.R + dy * kr1)
   c1.G = Fix(p1.Color.G + dy * kg1)
   c1.B = Fix(p1.Color.B + dy * kb1)
   c2.R = Fix(p2.Color.R + dy * kr2)
   c2.G = Fix(p2.Color.G + dy * kg2)
   c2.B = Fix(p2.Color.B + dy * kb2)
   If x1 = x2 Then
    If i < p.ScaleHeight And i >= 0 And j >= 0 And j < p.ScaleWidth Then
     If z1 - ZBuffer(i * p.ScaleWidth + x1) > AntiMix And z1 < Camera Then
      p.PSet (x1, i), RGB(c1.R, c1.G, c1.B)
      ZBuffer(i * p.ScaleWidth + x1) = z1
     End If
    End If
   Else
    krX = (c1.R - c2.R) / (x1 - x2)
    kgX = (c1.G - c2.G) / (x1 - x2)
    kbX = (c1.B - c2.B) / (x1 - x2)
    kzX = (z1 - z2) / (x1 - x2)
    For j = x1 To x2
     tr = Fix(c1.R + (j - x1) * krX)
     tg = Fix(c1.G + (j - x1) * kgX)
     tb = Fix(c1.B + (j - x1) * kbX)
     tz = z1 + (j - x1) * kzX
     If i < p.ScaleHeight And i >= 0 And j >= 0 And j < p.ScaleWidth Then
      If tz - ZBuffer(i * p.ScaleWidth + j) > AntiMix And tz < Camera Then
       p.PSet (j, i), RGB(tr, tg, tb)
       ZBuffer(i * p.ScaleWidth + j) = tz
      End If
     End If
    Next j
   End If
  Next i
 Else
  k1 = (p1.X - p2.X) / (p1.Y - p2.Y)
  k2 = (p1.X - p3.X) / (p1.Y - p3.Y)
  kz1 = (p1.Z - p2.Z) / (p1.Y - p2.Y)
  kz2 = (p1.Z - p3.Z) / (p1.Y - p3.Y)
  kr1 = (p1.Color.R - p2.Color.R) / (p1.Y - p2.Y)
  kg1 = (p1.Color.G - p2.Color.G) / (p1.Y - p2.Y)
  kb1 = (p1.Color.B - p2.Color.B) / (p1.Y - p2.Y)
  kr2 = (p1.Color.R - p3.Color.R) / (p1.Y - p3.Y)
  kg2 = (p1.Color.G - p3.Color.G) / (p1.Y - p3.Y)
  kb2 = (p1.Color.B - p3.Color.B) / (p1.Y - p3.Y)
  For i = p1.Y To p3.Y
   dy = i - p1.Y
   x1 = Fix(p1.X + dy * k1)
   x2 = Fix(p1.X + dy * k2)
   z1 = p1.Z + dy * kz1
   z2 = p1.Z + dy * kz2
   c1.R = Fix(p1.Color.R + dy * kr1)
   c1.G = Fix(p1.Color.G + dy * kg1)
   c1.B = Fix(p1.Color.B + dy * kb1)
   c2.R = Fix(p1.Color.R + dy * kr2)
   c2.G = Fix(p1.Color.G + dy * kg2)
   c2.B = Fix(p1.Color.B + dy * kb2)
   If x1 = x2 Then
    If i < p.ScaleHeight And i >= 0 And j >= 0 And j < p.ScaleWidth Then
     If z1 - ZBuffer(i * p.ScaleWidth + x1) > AntiMix And z1 < Camera Then
      p.PSet (x1, i), RGB(c1.R, c1.G, c1.B)
      ZBuffer(i * p.ScaleWidth + x1) = z1
     End If
    End If
   Else
    krX = (c1.R - c2.R) / (x1 - x2)
    kgX = (c1.G - c2.G) / (x1 - x2)
    kbX = (c1.B - c2.B) / (x1 - x2)
    kzX = (z1 - z2) / (x1 - x2)
    For j = x1 To x2
     tr = Fix(c1.R + (j - x1) * krX)
     tg = Fix(c1.G + (j - x1) * kgX)
     tb = Fix(c1.B + (j - x1) * kbX)
     tz = z1 + (j - x1) * kzX
     If i < p.ScaleHeight And i >= 0 And j >= 0 And j < p.ScaleWidth Then
      If tz - ZBuffer(i * p.ScaleWidth + j) > AntiMix And tz < Camera Then
       p.PSet (j, i), RGB(tr, tg, tb)
       ZBuffer(i * p.ScaleWidth + j) = tz
      End If
     End If
    Next j
   End If
  Next i
 End If
End Sub

Public Sub DrawTriangle(p As PictureBox, t As Triangle)
 Dim p1 As P2D, p2 As P2D, p3 As P2D, tmp As P2D, p2tmp As P2D
 Dim k As Double, kr As Double, kg As Double, kb As Double
 p1 = t.Point2D(0)
 p2 = t.Point2D(1)
 p3 = t.Point2D(2)
 If p1.Y > p2.Y Then
  tmp = p1
  p1 = p2
  p2 = tmp
 End If
 If p1.Y > p3.Y Then
  tmp = p1
  p1 = p3
  p3 = tmp
 End If
 If p2.Y > p3.Y Then
  tmp = p2
  p2 = p3
  p3 = tmp
 End If
 If p2.Y = p1.Y Then
  If p2.X < p1.X Then
   tmp = p2
   p2 = p1
   p1 = tmp
  End If
  DrawSimpleTriangle p, p1, p2, p3
  Exit Sub
 End If
 If p2.Y = p3.Y Then
  If p2.X > p3.X Then
   tmp = p3
   p3 = p2
   p2 = tmp
  End If
  DrawSimpleTriangle p, p1, p2, p3
  Exit Sub
 End If
 k = (p3.X - p1.X) / (p3.Y - p1.Y)
 kr = (p3.Color.R - p1.Color.R) / (p3.Y - p1.Y)
 kg = (p3.Color.G - p1.Color.G) / (p3.Y - p1.Y)
 kb = (p3.Color.B - p1.Color.B) / (p3.Y - p1.Y)
 kz = (p3.Z - p1.Z) / (p3.Y - p1.Y)
 p2tmp.Y = p2.Y
 dy = p2.Y - p1.Y
 p2tmp.X = p1.X + Fix(k * dy)
 p2tmp.Color.R = p1.Color.R + Fix(kr * dy)
 p2tmp.Color.G = p1.Color.G + Fix(kg * dy)
 p2tmp.Color.B = p1.Color.B + Fix(kb * dy)
 p2tmp.Z = p1.Z + Fix(kz * dy)
 If p2tmp.X < p2.X Then
  tmp = p2
  p2 = p2tmp
  p2tmp = tmp
 End If
 DrawSimpleTriangle p, p1, p2, p2tmp
 DrawSimpleTriangle p, p2, p2tmp, p3
End Sub

Public Sub DrawTriangleLine(p As PictureBox, t As Triangle)
 p.Line (t.Point2D(0).X, t.Point2D(0).Y)-(t.Point2D(1).X, t.Point2D(1).Y)
 p.Line (t.Point2D(1).X, t.Point2D(1).Y)-(t.Point2D(2).X, t.Point2D(2).Y)
 p.Line (t.Point2D(2).X, t.Point2D(2).Y)-(t.Point2D(0).X, t.Point2D(0).Y)
End Sub
