VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "打鴨子"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   FillColor       =   &H00FF00FF&
   FillStyle       =   2  '水平線
   Icon            =   "打鴨子.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12195
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer 烏鴉計時器 
      Interval        =   150
      Left            =   7080
      Top             =   2880
   End
   Begin VB.Timer 飛彈計時器 
      Interval        =   20
      Left            =   360
      Top             =   6960
   End
   Begin VB.Timer 鴨子計時器 
      Interval        =   200
      Left            =   600
      Top             =   960
   End
   Begin VB.Label 狀況 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7170
      TabIndex        =   6
      ToolTipText     =   "狀況"
      Top             =   7680
      Width           =   4155
   End
   Begin VB.Label 命中數量 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "命中數量"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label 命中數量標簽 
      BackColor       =   &H0000FFFF&
      Caption         =   "命中數量"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label 發射數量 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "發射數量"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label 發射數量標簽 
      BackColor       =   &H0000FFFF&
      Caption         =   "發射數量"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Image 飛彈 
      Enabled         =   0   'False
      Height          =   480
      Left            =   1320
      Picture         =   "打鴨子.frx":0442
      ToolTipText     =   "飛彈"
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label 班級座號 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '單線固定
      Caption         =   "表單:Form2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   3180
      TabIndex        =   1
      Top             =   120
      Width           =   2265
   End
   Begin VB.Line 底線 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   5
      X1              =   0
      X2              =   12000
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label 專案說明 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '單線固定
      Caption         =   "打鴨子遊戲"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6690
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.Line 頂線 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   5
      X1              =   0
      X2              =   12000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image 烏鴉 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   7080
      Picture         =   "打鴨子.frx":1104
      Stretch         =   -1  'True
      ToolTipText     =   "烏鴉"
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Image 鴨子 
      Enabled         =   0   'False
      Height          =   1065
      Left            =   1200
      Picture         =   "打鴨子.frx":5492
      Stretch         =   -1  'True
      ToolTipText     =   "鴨子"
      Top             =   840
      Width           =   1155
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***注意單引號後面是程式註解說明用，在真正考試寫程式是不必寫的

Dim 鴨子移動距離 As Integer     '宣告鴨子移動用的變數
Dim 烏鴉移動距離 As Integer     '宣告烏鴉移動用的變數
Dim 飛彈已發射 As Boolean       '宣告布林來記住飛彈<是,否>正往上飛

Private Sub Form_Load()            '程式一開始執行時，會先執行的事件，用來做變數初值設定
鴨子移動距離 = 300                 '先決定鴨子每次移動的距離，此值越大，移動越快
烏鴉移動距離 = 300                 '先決定烏鴉每次移動的距離，此值越大，移動越快
飛彈已發射 = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)                                    '當滑鼠在表單移動，就檢查飛彈
If 飛彈已發射 = False Then          '如果飛彈沒發射，就要跟著滑鼠左右移動
   飛彈.Left = X - 飛彈.Width / 2   '沒射出去的飛彈的左界保持跟滑鼠的X值相關，而飛彈上界不要跟，這樣飛彈就會左右受控，上下不變。
End If                              '這樣看起來，就像拿飛彈在瞄準鴨子
End Sub

Private Sub Form_Click()           '表單被點一下，代表玩的人要發射一個飛彈
If 飛彈已發射 = False Then      '如果飛彈還未射出
   飛彈已發射 = True            '就讓那個飛彈的狀態設為真，這樣飛彈就發射出去了
   發射數量 = 發射數量 + 1        '每發射了一個飛彈，當然發射數量得加一
   狀況 = "已發射...."
End If
End Sub

Private Sub 鴨子計時器_Timer()                  '固定時間週期內會來執行一遍的事件
鴨子.Left = 鴨子.Left + 鴨子移動距離            '鴨子每次往一個方向移動固定距離
If 鴨子.Left + 鴨子.Width > Me.Width Then    '檢查是否到達表單右界的外面
   鴨子移動距離 = -300                          '如果是，代表已加太多，把移動距離符號相反，改成負的
End If
If 鴨子.Left < 0 Then                           '檢查是否到達表單左界的左邊
   鴨子移動距離 = 300                           '如果是，代表已減太多(加太多負?)，把移動距離符號相反，改成正的
End If
End Sub
Private Sub 烏鴉計時器_Timer()                  '固定時間週期內會來執行一遍的事件
烏鴉.Left = 烏鴉.Left + 烏鴉移動距離            '烏鴉每次往一個方向移動固定距離
If 烏鴉.Left + 烏鴉.Width > Me.Width Then    '檢查是否到達表單右界的外面,Me是代表本表單
   烏鴉移動距離 = -300                          '如果是，代表已加太多，把移動距離符號相反，改成負
End If
If 烏鴉.Left < 0 Then                           '檢查是否到達表單左界的左邊
   烏鴉移動距離 = 300                           '如果是，代表已減太多(加太多負?)，把移動距離符號相反，改成正的
End If
End Sub



Private Sub 飛彈計時器_Timer()       '固定時間週期內會來執行一遍的事件
If 飛彈已發射 = True Then            '如果飛彈已發射，才須往上移動及比較，如果不是在這裡不必管它
   飛彈.Top = 飛彈.Top - 150         '每次往上移150點，可放此值，改變飛彈上昇的整度
   X = 飛彈.Left + 飛彈.Width / 2    '算出目前這個飛彈的砥火位置
   Y = 飛彈.Top + 200                '把飛彈化成一個X,Y作標，使圖形重疊的比對變得容易
   If Y < 頂線.Y1 Then               '如果飛彈上界小於頂線，代表飛到表單上面去了(注意電腦用第四象限)
      飛彈.Top = 底線.Y1 - 飛彈.Height - 60       '把該個飛彈收回到底線上方一點，準備下次發射。
      飛彈已發射 = False             '就把該枚飛彈的發射狀態設為假，使其不再飛行
      狀況 = "損失飛彈一枚"
   End If
   If X > 鴨子.Left And X < 鴨子.Left + 鴨子.Width And Y > 鴨子.Top And Y < 鴨子.Top + 鴨子.Height Then    '如果X在鴨子的左右界之間
      飛彈.Top = 底線.Y1 - 飛彈.Height - 60                         ''且Y在鴨子的上下界之間, 收回飛彈。
      飛彈已發射 = False                                            '打中數量加一，並如上面方式
      狀況 = "打中目標"
      命中數量 = 命中數量 + 1                  '代表圖形重疊，飛彈打中鴨子
   End If
   If X > 烏鴉.Left And X < 烏鴉.Left + 烏鴉.Width And Y > 烏鴉.Top And Y < 烏鴉.Top + 烏鴉.Height Then  '如果X在鴨子的左右界之間
      飛彈.Top = 底線.Y1 - 飛彈.Height - 60                         ''且Y在烏鴉的上下界之間,並如上面方式收回飛彈。
      飛彈已發射 = False                                            '代表圖形重疊，飛彈打中烏鴉
      狀況 = "打錯目標"
   End If
End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '當按了表單右上角X時會發生的事件
Form4.Show                              '因為是多重表單，所以在本表單結束前要啟動選項表單。
End Sub                                 '考試時此段不必寫

