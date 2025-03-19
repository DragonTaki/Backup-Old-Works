VERSION 5.00
Begin VB.Form AngleAndCoordinateTransform 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "Angle & Coordinate Transform"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   3990
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   4000
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Output 
      BackColor       =   &H00000000&
      Caption         =   "Output"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   2040
      TabIndex        =   8
      Top             =   240
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   14
         Top             =   2640
         Width           =   1575
      End
   End
   Begin VB.Frame Input 
      BackColor       =   &H00000000&
      Caption         =   "Input"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1935
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
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "Rec X"
         Top             =   480
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
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Rec Y"
         Top             =   1080
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
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Pol R"
         Top             =   1680
         Width           =   1575
      End
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
         TabIndex        =   1
         Text            =   "0"
         ToolTipText     =   "Pol Theta"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label RecX 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Rec X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Calculate 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Calculate"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label PolTheta 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Pol Theta"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label PolR 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Pol R"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label RecY 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "Rec Y"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Label WindowsVersionLabel 
      BackColor       =   &H00000000&
      Caption         =   "Running program in "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4095
   End
   Begin VB.Menu FileBaseMenu 
      Caption         =   "File(F&)"
      Begin VB.Menu CalculateMenu 
         Caption         =   "Calculate(C&)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ResetMenu 
         Caption         =   "Reset(R&)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu SeparatorMenu0 
         Caption         =   "-"
      End
      Begin VB.Menu EndMenu 
         Caption         =   "End"
      End
   End
   Begin VB.Menu InfMenu 
      Caption         =   "Inf(I&)"
      Begin VB.Menu AuthorMenu 
         Caption         =   "Author(A&)"
      End
   End
End
Attribute VB_Name = "AngleAndCoordinateTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pi As Double

Private Sub AuthorMenu_Click()
   Dim Temp As Integer
   
   Temp = MsgBox("Angle And Coordinate Transform " & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "Made By Synasaivaltos, synasaivaltos@gmail.com", vbOKOnly + vbInformation, "Angle And Coordinate Transform")
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
Private Sub CalculateMenu_Click()
   Call Calculate_Click
End Sub
Private Sub EndMenu_Click()
   End
End Sub
Private Sub Form_Load()
   Call AngleAndCoordinateTransformModule.GetWindowsVersion
   WindowsVersionLabel.Caption = WindowsVersionLabel.Caption & GetWindowsVersion & "."
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
Private Sub ResetMenu_Click()
   Call Reset_Click
End Sub
