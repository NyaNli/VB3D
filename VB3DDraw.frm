VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB3D"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   9855
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   3000
      TabIndex        =   15
      Text            =   "1"
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "真实渲染"
      Height          =   615
      Left            =   6720
      TabIndex        =   12
      Top             =   9960
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "线框渲染"
      Height          =   615
      Left            =   6720
      TabIndex        =   11
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "重置模型"
      Height          =   1455
      Left            =   8280
      TabIndex        =   10
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "应用设定"
      Height          =   1455
      Left            =   5160
      TabIndex        =   9
      Top             =   9120
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3000
      TabIndex        =   8
      Text            =   "4000"
      Top             =   9960
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   3000
      TabIndex        =   7
      Text            =   "45"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3000
      TabIndex        =   6
      Text            =   "250"
      Top             =   9240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   8760
      Width           =   9615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开文件"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   9120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   365
      ScaleLeft       =   200
      ScaleMode       =   0  'User
      ScaleTop        =   200
      ScaleWidth      =   437
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   0
         TabIndex        =   13
         Top             =   7440
         Width           =   2655
      End
   End
   Begin VB.Label Label5 
      Caption         =   "混合修正："
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   10320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "渲染距离："
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   9960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "FOV："
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "摄像机距离："
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   9240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Surface() As Triangle, NowSurface() As Triangle
Dim rX As Single, rY As Single, rotateX As Double, rotateY As Double
Dim offX As Double, offY As Double, SizeUp As Double
Dim RotateSpeed As Double
Dim Loaded As Boolean, Press As Boolean, LineMode As Boolean

Private Sub Render()
 If Not Loaded Then
  Exit Sub
 End If
 Draw_Clear Picture1
 For i = 0 To UBound(NowSurface)
  MPST Picture1, NowSurface(i), offX, offY, SizeUp
  If LineMode Then
   DrawTriangleLine Picture1, NowSurface(i)
  Else
   DrawTriangle Picture1, NowSurface(i)
  End If
 Next
End Sub

Private Sub ResetModel()
 rotateX = 0: rotateY = 0: offX = 0: offY = 0: SizeUp = 1
 UpdateText
 For i = 0 To UBound(NowSurface)
  For j = 0 To 2
   With NowSurface(i).Point3D(j)
    .X = Surface(i).Point3D(j).X
    .Y = Surface(i).Point3D(j).Y
    .Z = Surface(i).Point3D(j).Z
    .Color.R = Surface(i).Point3D(j).Color.R
    .Color.G = Surface(i).Point3D(j).Color.G
    .Color.B = Surface(i).Point3D(j).Color.B
   End With
  Next j
 Next i
End Sub

Private Sub UpdateText()
 If LineMode Then
 Label4.Caption = "[线框渲染]" & vbCrLf & "三角形数：" & UBound(Surface) + 1 & vbCrLf & "X轴旋转：" & rotateX & " : 偏移：" & offX & vbCrLf & "Y轴旋转：" & rotateY & " : 偏移：" & offY & vbCrLf & "缩放：" & SizeUp
 Else
 Label4.Caption = "[真实渲染]" & vbCrLf & "三角形数：" & UBound(Surface) + 1 & vbCrLf & "X轴旋转：" & rotateX & " : 偏移：" & offX & vbCrLf & "Y轴旋转：" & rotateY & " : 偏移：" & offY & vbCrLf & "缩放：" & SizeUp
 End If
End Sub

Private Sub Command1_Click()
 'CommonDialog1.ShowOpen
 If Text1.Text = "" Or Dir(Text1.Text) = "" Then
  Exit Sub
 End If
 Open Text1.Text For Random As #1 Len = 90
 n = LOF(1) / 90
 ReDim Surface(0 To n - 1), NowSurface(0 To n - 1)
 For i = 0 To n - 1
  Get #1, i + 1, Surface(i).Point3D
 Next i
 Close #1
 ResetModel
 Loaded = True
 Render
End Sub

Private Sub Command2_Click()
 If IsNumeric(Text2.Text) Then
  Camera = Val(Text2.Text)
 Else
  Text2.Text = Camera
 End If
 If IsNumeric(Text3.Text) Then
  FOV = Val(Text3.Text)
 Else
  Text3.Text = FOV
 End If
 If IsNumeric(Text4.Text) Then
  RenderDistance = Val(Text4.Text)
 Else
  Text4.Text = RenderDistance
 End If
 If IsNumeric(Text5.Text) Then
  AntiMix = Val(Text5.Text)
 Else
  Text5.Text = AntiMix
 End If
 Render
End Sub

Private Sub Command3_Click()
 If Not Loaded Then
  Exit Sub
 End If
 ResetModel
 Render
End Sub

Private Sub Command4_Click()
 LineMode = True
 UpdateText
 Render
End Sub

Private Sub Command5_Click()
 LineMode = False
 UpdateText
 Render
End Sub

Private Sub Form_Load()
 'Dim test As P3D
 'MsgBox Len(test)
 Loaded = False
 Press = False
 'CommonDialog1.Filter = "V3D模型文件(*.v3d)|*.v3d"
 Camera = 250
 FOV = 45
 RenderDistance = 4000
 AntiMix = 1
 RotateSpeed = 0.5
 Draw_Init Picture1
 rotateX = 0
 rotateY = 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Not Loaded Then
  Exit Sub
 End If
 If Button = 1 Or Button = 2 Or Button = 4 Then
  rX = X
  rY = Y
  Picture1.Cls
  For i = 0 To UBound(NowSurface)
   MPST Picture1, NowSurface(i), offX, offY, SizeUp
   DrawTriangleLine Picture1, NowSurface(i)
  Next
  Press = True
 End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim tox As Double, toy As Double, toz As Double
 If Not Press Then
  Exit Sub
 End If
 If Button = 1 Then
  rotateX = rotateX + RotateSpeed * (X - rX)
  rotateY = rotateY + RotateSpeed * (Y - rY)
  Picture1.Cls
  For i = 0 To UBound(NowSurface)
   For j = 0 To 2
    With NowSurface(i).Point3D(j)
     .X = Surface(i).Point3D(j).X
     .Y = Surface(i).Point3D(j).Y
     .Z = Surface(i).Point3D(j).Z
     Rotate2D .Z, .X, rotateX
     Rotate2D .Y, .Z, rotateY
    End With
   Next j
   MPST Picture1, NowSurface(i), offX, offY, SizeUp
   DrawTriangleLine Picture1, NowSurface(i)
  Next i
 ElseIf Button = 2 Then
  Picture1.Cls
  offX = offX + X - rX
  offY = offY - Y + rY
  For i = 0 To UBound(NowSurface)
   MPST Picture1, NowSurface(i), offX, offY, SizeUp
   DrawTriangleLine Picture1, NowSurface(i)
  Next i
 ElseIf Button = 4 Then
  Picture1.Cls
  SizeUp = SizeUp - 0.005 * (Y - rY)
  If SizeUp < 0.5 Then
   SizeUp = 0.5
  End If
  For i = 0 To UBound(NowSurface)
   MPST Picture1, NowSurface(i), offX, offY, SizeUp
   DrawTriangleLine Picture1, NowSurface(i)
  Next i
 End If
 rX = X
 rY = Y
 UpdateText
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Not Press Then
  Exit Sub
 End If
 If Button = 1 Or Button = 2 Or Button = 4 Then
  Render
  Press = False
 End If
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Text1.Text = Data.Files(1)
 Command1_Click
End Sub
