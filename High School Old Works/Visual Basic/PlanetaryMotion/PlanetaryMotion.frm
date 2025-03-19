VERSION 5.00
Begin VB.Form PlanetaryMotion 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "Planetary Motion"
   ClientHeight    =   8880
   ClientLeft      =   2400
   ClientTop       =   1650
   ClientWidth     =   10545
   ForeColor       =   &H00FFFFFF&
   Icon            =   "PlanetaryMotion.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10545
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Star 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4965
      TabIndex        =   17
      Top             =   4133
      Width           =   615
      Begin VB.Label StarLabel 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Star"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape StarShape 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         Shape           =   3  '圓形
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame AsteroidFrame 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Start 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         Style           =   1  '圖片外觀
         TabIndex        =   5
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox TimeMagnificationText 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  '暫止
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox AsteroidxText 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  '暫止
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox AsteroidyText 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  '暫止
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox AsteroidHeightText 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  '暫止
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox AsteroidWidthText 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  '暫止
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Reset 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reset"
         Height          =   375
         Left            =   600
         Style           =   1  '圖片外觀
         TabIndex        =   6
         Top             =   4560
         Width           =   495
      End
      Begin VB.CommandButton Conceal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hide"
         Height          =   375
         Left            =   1080
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label TimeMagnification 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Time Magnification"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Velocity 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00000000&
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Asteroidx 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Asteroid(x)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Asteroidy 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Asteroid(y)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label AsteroidHeight 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Asteroid Speed(x)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label AsteroidWidth 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Asteroid Speed(y)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label CurrentSpeed 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Current Speed"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   1455
      End
   End
   Begin VB.Timer PlanetTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   10080
      Top             =   0
   End
   Begin VB.Frame Asteroid 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4080
      TabIndex        =   16
      Top             =   960
      Width           =   615
      Begin VB.Label AsteroidLabel 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         BackStyle       =   0  '透明
         Caption         =   "Asteroid"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape AsteroidShape 
         BackColor       =   &H00808080&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         Shape           =   3  '圓形
         Top             =   0
         Width           =   615
      End
   End
End
Attribute VB_Name = "PlanetaryMotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dx(2) As Double
Dim Dy(2) As Double
Dim x As Double
Dim y As Double
Dim FormLoad As Boolean
Const AsteroidTop = 960
Const AsteroidLeft = 4080
Const AsteroidHeightSpeed = 185.00001
Const AsteroidWidthSpeed = -55.00001
Const PlanetTimerMagnification = 20
Const AsteroidFrameCaption = "Information"
Private Sub AsteroidxText_Change()
   If FormLoad = True Then
      If AsteroidxText.Text = "" Then
         Asteroid.Left = 0
      ElseIf Val(AsteroidxText.Text) > 0 Then
         If Val(AsteroidxText.Text) <= PlanetaryMotion.ScaleWidth Then
            Asteroid.Left = Val(AsteroidxText.Text)
         ElseIf Val(AsteroidxText.Text) > PlanetaryMotion.ScaleWidth Then
            AsteroidxText.Text = PlanetaryMotion.ScaleWidth
         End If
      Else
         Asteroid.Left = 0
         AsteroidxText.Text = ""
      End If
   End If
End Sub
Private Sub AsteroidyText_Change()
   If FormLoad = True Then
      If AsteroidyText.Text = "" Then
         Asteroid.Top = 0
      ElseIf Val(AsteroidyText.Text) > 0 Then
         If Val(AsteroidyText.Text) <= PlanetaryMotion.ScaleHeight Then
         Asteroid.Top = Val(AsteroidyText.Text)
         ElseIf Val(AsteroidyText.Text) > PlanetaryMotion.ScaleHeight Then
            AsteroidyText.Text = PlanetaryMotion.ScaleHeight
         End If
      Else
         Asteroid.Top = 0
         AsteroidyText.Text = ""
      End If
   End If
End Sub
Private Sub TimeMagnificationText_Change()
   If Val(TimeMagnificationText.Text) < 1 Or TimeMagnificationText.Text = "" Then
      TimeMagnificationText.Text = 1
   ElseIf Val(TimeMagnificationText.Text) <= 1000 And Val(TimeMagnificationText.Text) >= 1 Then
      PlanetTimer.Interval = Abs(Val(TimeMagnificationText.Text))
   Else
      TimeMagnificationText.Text = 1000
      PlanetTimer.Interval = 1000
   End If
End Sub
Private Sub Conceal_Click()
   If Conceal.Caption = "Hide" Then
      Conceal.Left = 120
      Conceal.Top = 240
      Conceal.Caption = "Unhide"
      AsteroidFrame.Caption = ""
      AsteroidFrame.Height = 735
      AsteroidFrame.Width = 975
      Conceal.Width = 735
      Asteroidx.Visible = False
      AsteroidxText.Visible = False
      Asteroidy.Visible = False
      AsteroidyText.Visible = False
      AsteroidHeight.Visible = False
      AsteroidHeightText.Visible = False
      AsteroidWidth.Visible = False
      AsteroidWidthText.Visible = False
      CurrentSpeed.Visible = False
      Velocity.Visible = False
      Start.Visible = False
      Reset.Visible = False
   Else
      Conceal.Caption = "Hide"
      Conceal.Left = 1080
      Conceal.Top = 4560
      Conceal.Width = 495
      AsteroidFrame.Caption = AsteroidFrameCaption
      AsteroidFrame.Height = 5055
      AsteroidFrame.Width = 1695
      Asteroidx.Visible = True
      AsteroidxText.Visible = True
      Asteroidy.Visible = True
      AsteroidyText.Visible = True
      AsteroidHeight.Visible = True
      AsteroidHeightText.Visible = True
      AsteroidWidth.Visible = True
      AsteroidWidthText.Visible = True
      CurrentSpeed.Visible = True
      Velocity.Visible = True
      Start.Visible = True
      Reset.Visible = True
   End If
End Sub
Private Sub Form_Load()
   AsteroidxText.Text = AsteroidLeft
   AsteroidyText.Text = AsteroidTop
   AsteroidHeightText.Text = FormatNumber(AsteroidHeightSpeed, 4)
   AsteroidWidthText.Text = FormatNumber(AsteroidWidthSpeed, 4)
   TimeMagnificationText.Text = PlanetTimerMagnification
   AsteroidFrame.Caption = AsteroidFrameCaption
   PlanetTimer.Interval = PlanetTimerMagnification
   FormLoad = True
End Sub
Private Sub PlanetTimer_Timer()
   x = Asteroid.Left - Star.Left
   y = Asteroid.Top - Star.Top
   Dx(1) = Asteroid.Left
   Dy(1) = Asteroid.Top
   Dx(0) = Dx(0) - 168000000 * x / (x * x + y * y) ^ 1.5
   Dy(0) = Dy(0) - 168000000 * y / (x * x + y * y) ^ 1.5
   Asteroid.Left = Asteroid.Left + Dx(0)
   Asteroid.Top = Asteroid.Top + Dy(0)
   PlanetaryMotion.Line (Dx(1), Dy(1))-(Asteroid.Left, Asteroid.Top), vbWhite
   Velocity.Caption = FormatNumber(Round((Dx(0) * Dx(0) + Dy(0) * Dy(0)) ^ 0.5 / 2, 2), 2) & " m/s"
   'Dx(1) = Dx(1) + 1800000 * x / (x * x + y * y) ^ 1.5
   'Dy(1) = Dy(1) + 1800000 * y / (x * x + y * y) ^ 1.5
   'Star.Left = Star.Left + dx(1)
   'Star.Top = Star.Top + dy(1)
   AsteroidxText.Text = Asteroid.Left
   AsteroidyText.Text = Asteroid.Top
   AsteroidHeightText.Text = FormatNumber(Dx(0), 4)
   AsteroidWidthText.Text = FormatNumber(Dy(0), 4)
   TimeMagnificationText.Text = PlanetTimer.Interval
   ''
   AsteroidRight = Asteroid.Left + Asteroid.Width
   AsteroidBotton = Asteroid.Top + Asteroid.Height
   StarRight = Star.Left + Star.Width
   StarBotton = Star.Top + Star.Height
   If Asteroid.Left < StarRight And AsteroidRight > Star.Left And Asteroid.Top < StarBotton And AsteroidBotton > Star.Top Then
      Temp = MsgBox("The Asteroid Is Crash!", vbOKOnly + vbExclamation, "Oh Shit")
      Call Reset_Click
   End If
End Sub
Private Sub Reset_Click()
   PlanetTimer.Enabled = False
   AsteroidxText.Text = AsteroidLeft
   AsteroidyText.Text = AsteroidTop
   AsteroidHeightText.Text = FormatNumber(AsteroidHeightSpeed, 4)
   AsteroidWidthText.Text = FormatNumber(AsteroidWidthSpeed, 4)
   Asteroid.Top = AsteroidTop
   Asteroid.Left = AsteroidLeft
   PlanetTimer.Interval = PlanetTimerMagnification
   Start.Caption = "Start"
   Velocity.Caption = ""
   PlanetaryMotion.Cls
End Sub
Private Sub Start_Click()
   If Start.Caption = "Start" Then
      'Asteroid.Top = AsteroidxText.Text
      'Asteroid.Left = AsteroidyText.Text
      Dx(0) = AsteroidHeightText.Text
      Dy(0) = AsteroidWidthText.Text
      'Dx(1) = -47.5
      'Dy(1) = 8
      PlanetTimer.Enabled = True
      Start.Caption = "Pause"
   Else
      PlanetTimer.Enabled = False
      Velocity.Caption = ""
      Start.Caption = "Start"
   End If
End Sub
