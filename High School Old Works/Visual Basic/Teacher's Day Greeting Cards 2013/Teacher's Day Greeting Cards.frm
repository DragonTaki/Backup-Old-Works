VERSION 5.00
Begin VB.Form GreetingCard 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "Teacher's Day Greeting Card"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9495
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   9000
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   8520
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8040
      Top             =   0
   End
   Begin VB.Frame GreetingCardFrame 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00000000&
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Image Image1 
         Height          =   435
         Left            =   1200
         Picture         =   "Teacher's Day Greeting Cards.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label MainText4 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "卡片製作人：徐旻仟、李威果"
         BeginProperty Font 
            Name            =   "華康儷金黑(P)"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   6360
         Width           =   4095
      End
      Begin VB.Label MainText3 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "新竹高中67屆17班敬上"
         BeginProperty Font 
            Name            =   "華康儷金黑(P)"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   495
         Left            =   840
         TabIndex        =   36
         Top             =   3600
         Width           =   4095
      End
      Begin VB.Label MainText1 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "感謝您在我們一年級時的教導"
         BeginProperty Font 
            Name            =   "華康儷金黑(P)"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   1440
         TabIndex        =   35
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label MainText2 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "祝你教師節快樂~"
         BeginProperty Font 
            Name            =   "華康儷金黑(P)"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   1440
         TabIndex        =   34
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Teacher 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "To               老師："
         BeginProperty Font 
            Name            =   "華康儷金黑(P)"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   495
         Left            =   600
         TabIndex        =   33
         Top             =   600
         Width           =   3375
      End
      Begin VB.Image GreetingCardImage 
         Height          =   6975
         Left            =   120
         Picture         =   "Teacher's Day Greeting Cards.frx":EA36
         Stretch         =   -1  'True
         Top             =   120
         Width           =   9495
      End
   End
   Begin VB.Frame MainFrame 
      Appearance      =   0  '平面
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00000000&
      Height          =   6975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Frame Answer 
         BackColor       =   &H00000000&
         Caption         =   "Answer"
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   2400
         TabIndex        =   24
         Top             =   4680
         Width           =   4335
         Begin VB.TextBox AnswerRecXText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            MaxLength       =   16
            TabIndex        =   38
            ToolTipText     =   "Answer Rec X"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox AnswerPolThetaText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   2640
            MaxLength       =   16
            TabIndex        =   27
            ToolTipText     =   "Answer Pol Theta"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox AnswerPolRText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   2640
            MaxLength       =   16
            TabIndex        =   26
            ToolTipText     =   "Answer Pol R"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox AnswerRecYText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            MaxLength       =   16
            TabIndex        =   25
            ToolTipText     =   "Answer Rec Y"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Done 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Done"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1440
            TabIndex        =   32
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label AnswerRecY 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Rec Y"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label AnswerPolR 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Pol R"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   30
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label AnswerPolTheta 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Pol Theta"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label AnswerRecX 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Rec X"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Output 
         BackColor       =   &H00000000&
         Caption         =   "Output"
         ForeColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   4800
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
         Begin VB.TextBox ResultPolRText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Result Pol R"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox ResultRecXText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   17
            ToolTipText     =   "Result Rec X"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox ResultPolThetaText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Result Pol Theta"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox ResultRecYText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   15
            ToolTipText     =   "Result Rec Y"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label ResultRecY 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Rec Y"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label ResultRecX 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Rec X"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label ResultPolTheta 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Pol Theta"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label ResultPolR 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Pol R"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Reset 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Reset"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2640
            Width           =   1575
         End
      End
      Begin VB.Frame Input 
         BackColor       =   &H00000000&
         Caption         =   "Input"
         ForeColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   2400
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
         Begin VB.TextBox PolThetaText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            TabIndex        =   8
            Text            =   "0"
            ToolTipText     =   "Pol Theta"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox PolRText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            TabIndex        =   7
            Text            =   "0"
            ToolTipText     =   "Pol R"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox RecYText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            TabIndex        =   6
            Text            =   "0"
            ToolTipText     =   "Rec Y"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox RecXText 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            IMEMode         =   3  '暫止
            Left            =   120
            TabIndex        =   5
            Text            =   "0"
            ToolTipText     =   "Rec X"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label RecY 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Rec Y"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label PolR 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Pol R"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label PolTheta 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Pol Theta"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Calculate 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Calculate"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label RecX 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00000000&
            Caption         =   "Rec X"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label InstructionLabel1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "請運用下列計算機換算Rec(3, 17)及Pol(3, 17) 並輸入下列欄位"
         BeginProperty Font 
            Name            =   "華康儷金黑(P)"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1335
         Left            =   2160
         TabIndex        =   3
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.Label ClickLabel 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "點一下以進入程式"
      BeginProperty Font 
         Name            =   "華康儷金黑(P)"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
End
Attribute VB_Name = "GreetingCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xx(200) As Integer
Dim Yy(200) As Integer
Dim vx(200) As Double
Dim vy(200) As Double
Dim vsb(200) As Integer
Dim g As Double
Dim i, j, Tmp As Integer
Dim mx, my As Integer
Dim Stage As Integer '0:Start 1:Can click 2:MainFrame
Dim Pi As Double
Private Sub Done_Click()
   If Stage = 2 Then
      If AnswerRecXText.Text = "2.86891426788911" And AnswerRecYText.Text = "0.87711511416821" And AnswerPolRText.Text = "17.2626765016321" And AnswerPolThetaText.Text = "79.9920201985587" Then
         MainFrame.Visible = False
         GreetingCardFrame.Visible = True
         Stage = 3
      Else
         Tmp = MsgBox("You have the wrong answer!", vbExclamation + vbOKOnly, "Error!")
      End If
   End If
End Sub
Private Sub Form_Click()
   If Stage = 1 Then
      Timer1.Enabled = False
      Timer2.Enabled = False
      MainFrame.Visible = True
      Stage = 2
   End If
End Sub
Private Sub Form_Load()
   Randomize
   '
   GreetingCard.BackColor = &H0&
   GreetingCard.Height = 7500
   GreetingCard.Width = 10000
   ClickLabel.Visible = False
   MainFrame.Height = GreetingCard.Height - 480
   MainFrame.Width = GreetingCard.Width
   MainFrame.Left = 0
   MainFrame.Top = 0
   MainFrame.BackColor = &H0&
   MainFrame.Visible = False
   GreetingCardFrame.Height = GreetingCard.Height - 480
   GreetingCardFrame.Width = GreetingCard.Width
   GreetingCardFrame.Left = 0
   GreetingCardFrame.Top = 0
   GreetingCardFrame.BackColor = &HFFFFFF
   GreetingCardFrame.Visible = False
   GreetingCardImage.Height = GreetingCardFrame.Height
   GreetingCardImage.Width = GreetingCardFrame.Width
   GreetingCardImage.Left = 0
   GreetingCardImage.Top = 0
   '
   For i = 0 To 199
     Xx(i) = 1
     Yy(i) = 1
     vx(i) = 1
     vy(i) = 1
     vsb(i) = 0
   Next
   g = 25
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mx = X
   my = Y
End Sub
Private Sub Timer1_Timer()
   Tmp = 0
   For i = 0 To 30
     For j = Tmp To 199
        If vsb(j) = 0 Then
          Tmp = j
          Exit For
        End If
     Next
     vsb(j) = 1
     Xx(j) = mx
     Yy(j) = my
     vx(j) = (250 + 20 * Rnd) * Sin(Rnd * 2 * 3.14159)
     vy(j) = (250 + 20 * Rnd) * Cos(Rnd * 2 * 3.14159)
   Next
End Sub
Private Sub Timer2_Timer()
   Tmp = 0
   For j = Tmp To 199
     If vsb(j) = 1 Then
        Tmp = j
        PSet (Xx(j), Yy(j)), vbBlack
        ''''''''''''''''''''''''''''''''改這裡（與背景同色）
        Xx(j) = Xx(j) + vx(j)
        Yy(j) = Yy(j) + vy(j)
        vy(j) = vy(j) + g
        If Xx(j) > 20000 Or Yy(j) > 10000 Or Xx(j) < -100 Or Yy(j) < -200 Then
           vsb(j) = 0
        Else
           PSet (Xx(j), Yy(j)), &HFFFFFF * Rnd
        End If
     End If
   Next
End Sub
Private Sub Timer3_Timer()
   Stage = 1
   ClickLabel.Visible = True
   Timer3.Enabled = False
End Sub
Private Sub Calculate_Click()
   Call PolToRec
   Call RecToPol
End Sub
Private Sub PolToRec()
   Dim r, Theta As Double
   
   Pi = 4 * Atn(1)
   r = Val(PolRText.Text)
   Theta = Val(PolThetaText.Text)
   ResultRecXText.Text = r * Cos(Theta * Pi / 180)
   ResultRecYText.Text = r * Sin(Theta * Pi / 180)
End Sub
Private Sub RecToPol()
   Dim X, Y As Double
   
   Pi = 4 * Atn(1)
   X = Val(RecXText.Text)
   Y = Val(RecYText.Text)
   ResultPolRText.Text = Sqr(X ^ 2 + Y ^ 2)
   If X <> 0 Then
      ResultPolThetaText.Text = Atn(Y / X) * 180 / Pi
   Else
      ResultPolThetaText.Text = 0
   End If
End Sub
Private Sub Reset_Click()
   RecXText.Text = 0
   RecYText.Text = 0
   PolRText.Text = 0
   PolThetaText.Text = 0
   ResultPolRText.Text = Empty
   ResultPolThetaText = Empty
   ResultRecXText = Empty
   ResultRecYText = Empty
End Sub
