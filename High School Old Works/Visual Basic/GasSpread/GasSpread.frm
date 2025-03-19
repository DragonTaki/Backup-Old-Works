VERSION 5.00
Begin VB.Form GasSpread 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GasSpread"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   FillColor       =   &H000000FF&
   ForeColor       =   &H00000000&
   Icon            =   "GasSpread.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   5415
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Timer Renew 
      Interval        =   100
      Left            =   4920
      Top             =   720
   End
   Begin VB.CommandButton Stop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop"
      Height          =   500
      Left            =   4320
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton Go 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start"
      Height          =   500
      Left            =   3240
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton StopAddGas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Draw Walls"
      Height          =   500
      Left            =   2160
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Width           =   1100
   End
   Begin VB.Timer Run 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   4440
      Top             =   720
   End
   Begin VB.CommandButton AddGas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Gas"
      Height          =   500
      Left            =   1080
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton Eraser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Erase All"
      Height          =   500
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   1100
   End
   Begin VB.Shape GasFrame 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   3135
      Left            =   0
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label Status 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status: Nothing"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   5415
   End
   Begin VB.Line L 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   1200
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "GasSpread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mouse_status As Integer
Dim i As Integer
Dim j As Integer
Dim Mx As Integer
Dim My As Integer
Dim Nl As Integer
Dim Nd As Integer
Dim Dx(10000) As Integer
Dim Dy(10000) As Integer
Dim Vx(10000) As Integer
Dim Vy(10000) As Integer
Dim Str As String
Private Function Ckfunction(ByVal A As Integer, ByVal B As Integer)
   PSet (A, B), vbBlue
      Dx(Nd) = A
      Dy(Nd) = B
      Vx(Nd) = Rnd() * 10 - 5
      Vy(Nd) = Rnd() * 10 - 5
      Nd = Nd + 1
End Function
Private Sub AddGas_Click()
   Mouse_status = 2
End Sub
Private Sub Eraser_Click()
   For i = 1 To Nl - 1
      Unload L(i)
   Next
   For i = 0 To Nd - 1
      PSet (Dx(i), Dy(i)), vbWhite
      Dx(i) = -10
      Dy(i) = -10
      Vx(i) = -100
      Vy(i) = -100
   Next
   Nl = 1
   Nd = 0
End Sub
Private Sub Form_Click()
   If Mouse_status = 2 Then
      For i = 0 To 10
         Call Ckfunction(Int(Mx + Rnd * 400 - 20), Int(My + Rnd * 400 - 200))
      Next
   End If
End Sub
Private Sub Form_Load()
   Randomize
   Mouse_status = 0
   Nl = 1
   Nd = 0
   For i = 0 To 2000
      Dx(i) = -10
      Dy(i) = -10
      Vx(i) = -100
      Vy(i) = -100
   Next
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Mouse_status = 0 Then
      Mouse_status = 1
      Load L(Nl)
      L(Nl).Visible = True
      L(Nl).X1 = x
      L(Nl).Y1 = y
      L(Nl).X2 = x
      L(Nl).Y2 = y
      Nl = Nl + 1
   End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Mx = x
   My = y
   If Mouse_status = 1 Then
      Load L(Nl)
      L(Nl).Visible = True
      L(Nl).X1 = L(Nl - 1).X2
      L(Nl).Y1 = L(Nl - 1).Y2
      L(Nl).X2 = x
      L(Nl).Y2 = y
      Nl = Nl + 1
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Mouse_status = 1 Then
      Mouse_status = 0
   End If
End Sub
Private Sub Form_Resize()
   If GasSpread.Width < 5535 Then
      If GasSpread.WindowState = 0 Then
         GasSpread.Width = 5535
      End If
   End If
   If GasSpread.Height < 4695 Then
      If GasSpread.WindowState = 0 Then
         GasSpread.Height = 4695
      End If
   End If
   If GasSpread.Width <> 2400 And GasSpread.Height <> 405 Then
      GasFrame.Width = GasSpread.ScaleWidth
      GasFrame.Height = GasSpread.ScaleHeight - Eraser.Height - Status.Height
      GasFrame.Top = Eraser.Top + Eraser.Height
      Status.Width = GasSpread.ScaleWidth
      Status.Top = GasFrame.Top + GasFrame.Height
   End If
End Sub
Private Sub Go_Click()
   Run.Enabled = True
End Sub
Private Sub Renew_Timer()
   If Mouse_status = 0 Then
      Str = "Nothing"
   ElseIf Mouse_status = 1 Then
      Str = "Drawing Lines"
   ElseIf Mouse_status = 2 Then
      Str = "Making Molecules"
   End If
   Status.Caption = "Status: " & Str
   For i = 0 To Nd - 1
      If Dx(i) > 10000 Or Dx(i) < -10000 Or Dy(i) > 10000 Or Dy(i) < -10000 Then
         Vx(i) = 0
         Vy(i) = 0
      End If
   Next
End Sub
Private Sub Run_Timer()
   For i = 0 To Nd - 1
      If Point(Dx(i), Dy(i) + 7) = vbRed Then
         Vy(i) = Vy(i) * -1 + Rnd * 2 - 1
      End If
      If Point(Dx(i), Dy(i) - 7) = vbRed Then
         Vy(i) = Vy(i) * -1 + Rnd * 2 - 1
      End If
      If Point(Dx(i) + 7, Dy(i)) = vbRed Then
         Vx(i) = Vx(i) * -1 + Rnd * 2 - 1
      End If
      If Point(Dx(i) - 7, Dy(i)) = vbRed Then
         Vx(i) = Vx(i) * -1 + Rnd * 2 - 1
      End If
      PSet (Dx(i), Dy(i)), vbWhite
      Dx(i) = Dx(i) + Vx(i)
      Dy(i) = Dy(i) + Vy(i)
      PSet (Dx(i), Dy(i)), vbBlue
   Next
End Sub
Private Sub StopAddGas_Click()
   Mouse_status = 0
End Sub
Private Sub Stop_Click()
   Run.Enabled = False
End Sub
