VERSION 5.00
Begin VB.Form 象棋 
   BackColor       =   &H005FA76F&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13050
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame 持棋 
      BackColor       =   &H005FA76F&
      Caption         =   "持棋"
      BeginProperty Font 
         Name            =   "華康海報體W12"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1935
      Left            =   4680
      TabIndex        =   5
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Frame 剩餘棋數 
      BackColor       =   &H005FA76F&
      Caption         =   "剩餘棋數"
      BeginProperty Font 
         Name            =   "華康海報體W12"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1935
      Left            =   480
      TabIndex        =   4
      Top             =   7560
      Width           =   3855
   End
   Begin VB.Image 換玩家二圖 
      Height          =   570
      Left            =   7129
      Picture         =   "M04_象棋0.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image 換玩家一圖 
      Height          =   570
      Left            =   1129
      Picture         =   "M04_象棋0.frx":0F5B
      Stretch         =   -1  'True
      Top             =   840
      Width           =   555
   End
   Begin VB.Label 玩家二姓名 
      Alignment       =   2  '置中對齊
      BackColor       =   &H005FA76F&
      BeginProperty Font 
         Name            =   "華康娃娃體(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   9747
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label 玩家一姓名 
      Alignment       =   2  '置中對齊
      BackColor       =   &H005FA76F&
      BeginProperty Font 
         Name            =   "華康娃娃體(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3754
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label 玩家二 
      Alignment       =   2  '置中對齊
      BackColor       =   &H005FA76F&
      Caption         =   "玩家二"
      BeginProperty Font 
         Name            =   "華康粗黑體(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7707
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label 玩家一 
      Alignment       =   2  '置中對齊
      BackColor       =   &H005FA76F&
      Caption         =   "玩家一"
      BeginProperty Font 
         Name            =   "華康粗黑體(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1714
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   11805
      X2              =   11805
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   10485
      X2              =   10485
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   9165
      X2              =   9165
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   7845
      X2              =   7845
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   6525
      X2              =   6525
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   5205
      X2              =   5205
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   3885
      X2              =   3885
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   2565
      X2              =   2565
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   1245
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   0
      Left            =   1365
      Picture         =   "M04_象棋0.frx":1EB6
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   1
      Left            =   2685
      Picture         =   "M04_象棋0.frx":2019
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   16
      Left            =   1365
      Picture         =   "M04_象棋0.frx":217C
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   8
      Left            =   1365
      Picture         =   "M04_象棋0.frx":22DF
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   9
      Left            =   2685
      Picture         =   "M04_象棋0.frx":2442
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   2
      Left            =   4005
      Picture         =   "M04_象棋0.frx":25A5
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   3
      Left            =   5325
      Picture         =   "M04_象棋0.frx":2708
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   4
      Left            =   6645
      Picture         =   "M04_象棋0.frx":286B
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   5
      Left            =   7965
      Picture         =   "M04_象棋0.frx":29CE
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   6
      Left            =   9285
      Picture         =   "M04_象棋0.frx":2B31
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   7
      Left            =   10605
      Picture         =   "M04_象棋0.frx":2C94
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   15
      Left            =   10605
      Picture         =   "M04_象棋0.frx":2DF7
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   14
      Left            =   9285
      Picture         =   "M04_象棋0.frx":2F5A
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   13
      Left            =   7965
      Picture         =   "M04_象棋0.frx":30BD
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   12
      Left            =   6645
      Picture         =   "M04_象棋0.frx":3220
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   11
      Left            =   5325
      Picture         =   "M04_象棋0.frx":3383
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   10
      Left            =   4005
      Picture         =   "M04_象棋0.frx":34E6
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   23
      Left            =   10605
      Picture         =   "M04_象棋0.frx":3649
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   22
      Left            =   9285
      Picture         =   "M04_象棋0.frx":37AC
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   21
      Left            =   7965
      Picture         =   "M04_象棋0.frx":390F
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   20
      Left            =   6645
      Picture         =   "M04_象棋0.frx":3A72
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   19
      Left            =   5325
      Picture         =   "M04_象棋0.frx":3BD5
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   18
      Left            =   4005
      Picture         =   "M04_象棋0.frx":3D38
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   17
      Left            =   2685
      Picture         =   "M04_象棋0.frx":3E9B
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   31
      Left            =   10605
      Picture         =   "M04_象棋0.frx":3FFE
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   30
      Left            =   9285
      Picture         =   "M04_象棋0.frx":4161
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   29
      Left            =   7965
      Picture         =   "M04_象棋0.frx":42C4
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   28
      Left            =   6645
      Picture         =   "M04_象棋0.frx":4427
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   27
      Left            =   5325
      Picture         =   "M04_象棋0.frx":458A
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   26
      Left            =   4005
      Picture         =   "M04_象棋0.frx":46ED
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   25
      Left            =   2685
      Picture         =   "M04_象棋0.frx":4850
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image 棋 
      Height          =   1215
      Index           =   24
      Left            =   1365
      Picture         =   "M04_象棋0.frx":49B3
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
End
Attribute VB_Name = "象棋"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 已按第一顆棋 As Boolean
Dim 已開局 As Boolean '選顏色用
Dim 棋已翻開(32) As Boolean '用位子記
Dim 空格(32) As Boolean '用位子記
Dim 現在換玩家二 As Boolean
Dim 玩家一是黑的 As Boolean
Dim 第一顆是自己的棋 As Boolean
Dim 第二顆是自己的棋 As Boolean
Dim 是可移動的 As Boolean
Dim 第一顆棋 As Integer
Dim 第二顆棋 As Integer
Dim 第一顆棋位置 As Integer
Dim 第二顆棋位置 As Integer
Dim 中間有棋 As Integer
Private Sub Form_Load()
   玩家一姓名.Caption = 主頁.Name1P
   玩家二姓名.Caption = 主頁.Name2P
   Randomize
   Call 洗牌
End Sub
Private Sub 洗牌()
   Dim 棋已發過(32) As Boolean
   For i = 0 To 31
      棋已發過(i) = False
   Next
   For i = 0 To 31
重選棋:
      要放的棋 = Int(Rnd * 32)
      If 棋已發過(要放的棋) = True Then
         GoTo 重選棋
      End If
      棋已發過(要放的棋) = True
      棋(i).Tag = 要放的棋
   Next
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Ans = MsgBox("你確定要結束？", vbYesNo, "")
   If Ans = vbYes Then
      主頁.Show
      象棋.Hide
      Name1P = Empty
      Name2P = Empty
   Else
      Cancel = 1
   End If
End Sub
Private Sub 棋_Click(Index As Integer)
   If 已開局 = False Then '決定顏色
      If 棋(Index).Tag = 0 Or 棋(Index).Tag = 1 Or 棋(Index).Tag = 2 Or 棋(Index).Tag = 3 Or 棋(Index).Tag = 4 Or 棋(Index).Tag = 5 Or 棋(Index).Tag = 6 Or 棋(Index).Tag = 7 Or 棋(Index).Tag = 8 Or 棋(Index).Tag = 9 Or 棋(Index).Tag = 10 Or 棋(Index).Tag = 11 Or 棋(Index).Tag = 12 Or 棋(Index).Tag = 13 Or 棋(Index).Tag = 14 Or 棋(Index).Tag = 15 Then '點到黑棋
         玩家一.ForeColor = vbBlack
         玩家二.ForeColor = vbRed
         玩家一是黑的 = True
      Else '點到紅棋
         玩家一.ForeColor = vbRed
         玩家二.ForeColor = vbBlack
      End If
      Call 換人
      已開局 = True
      棋(Index).Picture = LoadPicture(App.Path & "\" & 棋(Index).Tag & ".gif")
      棋已翻開(Index) = True
   Else '已決定顏色
      If 棋已翻開(Index) = False Then '該棋未翻開
         棋(Index).Picture = LoadPicture(App.Path & "\" & 棋(Index).Tag & ".gif")
         棋已翻開(Index) = True
         Call 換人
      ElseIf 已按第一顆棋 = False Then '該棋已翻開，是第一顆
         已按第一顆棋 = True
         第一顆棋 = 棋(Index).Tag
         第一顆棋位置 = Index
         Call 判斷第一顆是否為自己的棋
         If 第一顆是自己的棋 = False Then '不是自己的棋
            已按第一顆棋 = False '還沒按
            第一顆棋 = Empty
            第一顆棋位置 = Empty
         Else '是自己的棋
            第一顆是自己的棋 = False
         End If
      ElseIf 已按第一顆棋 = True And 棋已翻開(Index) = False Then '第二顆棋未翻開
         If 是否暗吃 = True Then '可暗吃
            GoTo 可暗吃 '視同已翻開
         End If
         '還沒按
      ElseIf 已按第一顆棋 = True And 棋已翻開(Index) = True Then '該棋已翻開，是第二顆
可暗吃:
         第二顆棋 = 棋(Index).Tag
         第二顆棋位置 = Index
         Call 判斷第二顆是否為自己的棋
         If 第二顆是自己的棋 = True Then
            第二顆是自己的棋 = False
            第二顆棋 = Empty
            第二顆棋位置 = Empty
            Exit Sub '第二顆是自己的棋，不能吃，離開
         Else
            '第二顆不是自己的棋，可以吃
         End If
         Call 檢查兩顆棋的位置
         If 是可移動的 = False Then
            Exit Sub
         End If
         是可移動的 = False
         中間有棋 = 0
         Call 動棋子
         已按第一顆棋 = False
         第一顆棋 = Empty
         第二顆棋 = Empty
         第一顆棋位置 = Empty
         第二顆棋位置 = Empty
         Call 換人
      End If
   End If
End Sub
Private Sub 判斷第一顆是否為自己的棋()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 第一顆棋 = 0 Or 第一顆棋 = 1 Or 第一顆棋 = 2 Or 第一顆棋 = 3 Or 第一顆棋 = 4 Or 第一顆棋 = 5 Or 第一顆棋 = 6 Or 第一顆棋 = 7 Or 第一顆棋 = 8 Or 第一顆棋 = 9 Or 第一顆棋 = 10 Or 第一顆棋 = 11 Or 第一顆棋 = 12 Or 第一顆棋 = 13 Or 第一顆棋 = 14 Or 第一顆棋 = 15 Then '第一顆是黑棋
            第一顆是自己的棋 = True
         End If
      Else '玩家一是紅的
         If 第一顆棋 = 16 Or 第一顆棋 = 17 Or 第一顆棋 = 18 Or 第一顆棋 = 19 Or 第一顆棋 = 20 Or 第一顆棋 = 21 Or 第一顆棋 = 22 Or 第一顆棋 = 23 Or 第一顆棋 = 24 Or 第一顆棋 = 25 Or 第一顆棋 = 26 Or 第一顆棋 = 27 Or 第一顆棋 = 28 Or 第一顆棋 = 29 Or 第一顆棋 = 30 Or 第一顆棋 = 31 Then '第一顆是紅棋
            第一顆是自己的棋 = True
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 第一顆棋 = 16 Or 第一顆棋 = 17 Or 第一顆棋 = 18 Or 第一顆棋 = 19 Or 第一顆棋 = 20 Or 第一顆棋 = 21 Or 第一顆棋 = 22 Or 第一顆棋 = 23 Or 第一顆棋 = 24 Or 第一顆棋 = 25 Or 第一顆棋 = 26 Or 第一顆棋 = 27 Or 第一顆棋 = 28 Or 第一顆棋 = 29 Or 第一顆棋 = 30 Or 第一顆棋 = 31 Then '第一顆是紅棋
            第一顆是自己的棋 = True
         End If
      Else '玩家二是黑的
         If 第一顆棋 = 0 Or 第一顆棋 = 1 Or 第一顆棋 = 2 Or 第一顆棋 = 3 Or 第一顆棋 = 4 Or 第一顆棋 = 5 Or 第一顆棋 = 6 Or 第一顆棋 = 7 Or 第一顆棋 = 8 Or 第一顆棋 = 9 Or 第一顆棋 = 10 Or 第一顆棋 = 11 Or 第一顆棋 = 12 Or 第一顆棋 = 13 Or 第一顆棋 = 14 Or 第一顆棋 = 15 Then '第一顆是黑棋
            第一顆是自己的棋 = True
         End If
      End If
   End If
End Sub
Private Sub 判斷第二顆是否為自己的棋()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '第一顆是黑棋
            第二顆是自己的棋 = True
         End If
      Else '玩家一是紅的
         If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then '第一顆是紅棋
            第二顆是自己的棋 = True
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 第一顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then '第一顆是紅棋
            第二顆是自己的棋 = True
         End If
      Else '玩家二是黑的
         If 第一顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '第一顆是黑棋
            第二顆是自己的棋 = True
         End If
      End If
   End If
End Sub
Private Sub 換人()
      If 現在換玩家二 = False Then
         現在換玩家二 = True
         換玩家一圖.Visible = False
         換玩家二圖.Visible = True
      Else
         現在換玩家二 = False
         換玩家一圖.Visible = True
         換玩家二圖.Visible = False
      End If
End Sub
Private Sub 檢查兩顆棋的位置()
   If 是否炮跳 = False Then '炮不跳
      GoTo 不開炮跳 '和一般棋動法一樣
   End If
   If 第一顆棋 <> 9 And 第一顆棋 <> 10 And 第一顆棋 <> 25 And 第一顆棋 <> 26 Then '第一顆棋不是包或炮
不開炮跳:
      If 第一顆棋位置 Mod 8 = 第二顆棋位置 Mod 8 And 第一顆棋位置 = 第二顆棋位置 + 8 Then '第二顆棋在第一顆棋上方
         是可移動的 = True
      ElseIf 第一顆棋位置 Mod 8 = 第二顆棋位置 Mod 8 And 第一顆棋位置 = 第二顆棋位置 - 8 Then '第二顆棋在第一顆棋下方
         是可移動的 = True
      ElseIf 第一顆棋位置 = 第二顆棋位置 + 1 Then '第二顆棋在第一顆棋右邊
         是可移動的 = True
      ElseIf 第一顆棋位置 = 第二顆棋位置 - 1 Then '第二顆棋在第一顆棋左邊
         是可移動的 = True
      End If
   Else '第一顆棋是包或炮
      If Abs((第一顆棋位置 - 第二顆棋位置)) < 8 Then '第一顆棋和第二顆棋同行
         If 棋(第二顆棋位置).Tag = -1 Then '第二顆棋是空的
            If 第一顆棋位置 = 第二顆棋位置 + 8 Or 第一顆棋位置 = 第二顆棋位置 - 8 Then '第二顆棋是空的且在第一顆棋上方或下方
               是可移動的 = True
            End If
         Else '第二顆棋不是空的
            temp1 = 第一顆棋位置
            temp2 = 第二顆棋位置
            If temp1 > temp2 Then
               temp1 = 第二顆棋位置
               temp2 = 第一顆棋位置
            End If
            For i = temp1 + 1 To temp2
               If 棋(第二顆棋位置).Tag <> -1 Then
                  中間有棋 = 中間有棋 + 1
               End If
            Next
            If 中間有棋 = 2 Then
               是可移動的 = True
            End If
         End If
      ElseIf (第一顆棋位置 - 第二顆棋位置) Mod 8 = 0 Then  '第一顆棋和第二顆棋同列
         If 棋(第二顆棋位置).Tag = -1 Then '第二顆棋是空的
            If 第一顆棋位置 = 第二顆棋位置 + 1 Or 第一顆棋位置 = 第二顆棋位置 - 1 Then '第二顆棋是空的且在第一顆棋左邊或右邊
               是可移動的 = True
            End If
         Else '第二顆棋不是空的
            temp1 = 第一顆棋位置
            temp2 = 第二顆棋位置
            If temp1 > temp2 Then
               temp1 = 第二顆棋位置
               temp2 = 第一顆棋位置
            End If
            For i = temp1 + 1 To temp2 Step 8
               If 棋(第二顆棋位置).Tag <> -1 Then
                  中間有棋 = 中間有棋 + 1
               End If
            Next
            If 中間有棋 = 2 Then
               是可移動的 = True
            End If
         End If
      End If
   End If
End Sub
Private Sub 動棋子()
   If 現在換玩家二 = False Then
      If 玩家一是黑的 = True Then
         Call 動黑棋
      Else '玩家一是紅的
         Call 動紅棋
      End If
   Else    '現在換玩家二
      If 玩家一是黑的 = True Then
         Call 動紅棋
      Else '玩家二是黑的
         Call 動黑棋
      End If
   End If
End Sub
Private Sub 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag()
   棋(第二顆棋位置).Tag = 棋(第一顆棋位置).Tag
   棋(第一顆棋位置).Tag = -1
End Sub
Private Sub 動黑棋()
   If 第一顆棋 = 0 Then '將
      If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then '將吃兵
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 1 Or 第一顆棋 = 2 Then '士
      If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 16 Then  '士吃帥
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 3 Or 第一顆棋 = 4 Then '象
      If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Then  '象吃帥仕
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 5 Or 第一顆棋 = 6 Then '車
      If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then   '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Then  '車吃帥仕相
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 7 Or 第一顆棋 = 8 Then '馬
      If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then   '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Then  '馬吃帥仕相硨
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 9 Or 第一顆棋 = 10 Then '包
      If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃同色
         ''
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then   '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 11 Or 第一顆棋 = 12 Or 第一顆棋 = 13 Or 第一顆棋 = 14 Or 第一顆棋 = 15 Then '卒
      If 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Then '吃同色或炮
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 16 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then   '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Then    '卒吃仕相硨碼
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
End Sub
Private Sub 動紅棋()
   If 第一顆棋 = 16 Then '帥
      If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '帥吃卒
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 17 Or 第一顆棋 = 18 Then '仕
      If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 0 Then '仕吃將
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 19 Or 第一顆棋 = 20 Then '相
      If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Then '相吃將士
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 21 Or 第一顆棋 = 22 Then '硨
      If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Then '相吃將士象
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 23 Or 第一顆棋 = 24 Then '碼
      If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃同色
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Then '相吃將士象車
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 25 Or 第一顆棋 = 26 Then '炮
      If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then  '吃同色
         ''
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
   If 第一顆棋 = 27 Or 第一顆棋 = 28 Or 第一顆棋 = 29 Or 第一顆棋 = 30 Or 第一顆棋 = 31 Then '兵
      If 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Or 第二顆棋 = 19 Or 第二顆棋 = 20 Or 第二顆棋 = 21 Or 第二顆棋 = 22 Or 第二顆棋 = 23 Or 第二顆棋 = 24 Or 第二顆棋 = 25 Or 第二顆棋 = 26 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Then  '吃同色或包
         ''
      ElseIf 空格(第二顆棋位置) = True Or 第二顆棋 = 0 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".gif")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      ElseIf 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Then  '兵吃士象車馬
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
      End If
   End If
End Sub

