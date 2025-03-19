VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "打地鼠"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   Icon            =   "打地鼠.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   6435
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer 倒數計時器 
      Interval        =   1000
      Left            =   5880
      Top             =   1920
   End
   Begin VB.Timer 地鼠計時器 
      Interval        =   600
      Left            =   480
      Top             =   4320
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "打中率"
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
      Height          =   435
      Left            =   4200
      TabIndex        =   9
      Top             =   4920
      Width           =   1890
   End
   Begin VB.Label 打中率 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '單線固定
      Caption         =   "%"
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
      Left            =   4140
      TabIndex        =   8
      ToolTipText     =   "最高紀錄"
      Top             =   5400
      Width           =   1920
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   8
      Left            =   3000
      Picture         =   "打地鼠.frx":0442
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   8
      Left            =   2880
      Shape           =   2  '橢圓形
      Top             =   3960
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   7
      Left            =   1920
      Picture         =   "打地鼠.frx":0D8B
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   7
      Left            =   1800
      Shape           =   2  '橢圓形
      Top             =   3960
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   6
      Left            =   840
      Picture         =   "打地鼠.frx":16D4
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   6
      Left            =   720
      Shape           =   2  '橢圓形
      Top             =   3960
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   5
      Left            =   3000
      Picture         =   "打地鼠.frx":201D
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   5
      Left            =   2880
      Shape           =   2  '橢圓形
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   4
      Left            =   1920
      Picture         =   "打地鼠.frx":2966
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   4
      Left            =   1800
      Shape           =   2  '橢圓形
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   3
      Left            =   840
      Picture         =   "打地鼠.frx":32AF
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   3
      Left            =   720
      Shape           =   2  '橢圓形
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   2
      Left            =   3000
      Picture         =   "打地鼠.frx":3BF8
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   2
      Left            =   2880
      Shape           =   2  '橢圓形
      Top             =   2160
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   1
      Left            =   1920
      Picture         =   "打地鼠.frx":4541
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   1
      Left            =   1800
      Shape           =   2  '橢圓形
      Top             =   2160
      Width           =   975
   End
   Begin VB.Image 地鼠 
      Height          =   735
      Index           =   0
      Left            =   840
      Picture         =   "打地鼠.frx":4E8A
      Stretch         =   -1  'True
      ToolTipText     =   "地鼠"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label 剩餘時間標簽 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "剩餘時間"
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
      Height          =   435
      Left            =   4200
      TabIndex        =   7
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Label 剩餘時間 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '單線固定
      Caption         =   "5"
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
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "剩餘時間"
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label 打中次數 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
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
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "打中次數"
      Top             =   4320
      Width           =   900
   End
   Begin VB.Label 已打次數 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
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
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "已打次數"
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label 打中次數標簽 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "打中次數"
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
      Height          =   435
      Left            =   4200
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label 已打次數標簽 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "已打次數"
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
      Height          =   435
      Left            =   4200
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label 專案說明 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '單線固定
      Caption         =   "打地鼠"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label 班級座號 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '單線固定
      Caption         =   "表單:Form1"
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
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      Height          =   375
      Index           =   0
      Left            =   720
      Shape           =   2  '橢圓形
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***注意單引號後面是程式註解說明用，在真正考試寫程式是不必寫的
'本表單名稱是Form1都採用簡易寫法，所以要先在<已打次數>和<打中次數>的Captione填上0，在<剩餘時間>的Caption填上5

Private Sub Form_Load()   '程式一開始執行時，會先執行的事件，用來做變數初值設定
Randomize                 '因為有使用到<Rnd>這函數，所以要加這行，使其亂數完全沒有規律性
End Sub

Private Sub 地鼠計時器_Timer()   '記得在Interval這個屬性上填300，代表每0.3秒跳動一次，固定時間週期內會來執行一遍的事件
For i = 0 To 8                   '有9隻地鼠
   地鼠(i).Visible = False       '先把9隻地鼠全部消失
Next                             '因為亂數函數只能產生0.0000000-0.9999999之間的數
抽到的號碼 = Int(Rnd * 9)              '有9隻地鼠，所以亂數函數乘9，不要有小數所以取Int
地鼠(抽到的號碼).Visible = True        '抽中的那隻老鼠，就讓他出現
End Sub

Private Sub Form_Click()              '玩者用滑鼠點表單
已打次數 = 已打次數 + 1                 '點中表單，代表沒打中地鼠，所以只要加<已打次數>就可以了
打中率 = Round((打中次數 / 已打次數) * 100, 2) & "%" ''只要<已打次數>有變化就得重算打中率,乘100是代表打中率是百分之多少，Round(??,2),其中2代表取小數2位的精確度。
End Sub

Private Sub 地鼠_Click(Index As Integer)  '當圖片被打到，就代表打中地鼠
已打次數 = 已打次數 + 1   '打中，沒打中，<已打次數>當然都得加一
打中次數 = 打中次數 + 1   '即然打中了，當然<打中次數>也要加一
打中率 = Round((打中次數 / 已打次數) * 100, 2) & "%" '只要<已打次數>有變化就得重算打中率,乘100是代表打中率是百分之多少，Round(??,2),其中2代表取小數2位的精確度。
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)   '當按了表單右上角X時會發生的事件'考試時不必寫此段程式
Form4.Show                '因為是多重表單，所以在本表單結束前要啟動選項表單。考試時不必寫
End Sub                   '考試時此段不必寫

Private Sub 倒數計時器_Timer()   '記得在Interval這個屬性上填1000，代表每1秒執行本程式一次'固定時間週期內會來執行一遍的事件
剩餘時間 = 剩餘時間 - 1         '每秒會執行一次，把剩下時間減一
If 剩餘時間 <= 0 Then            '每次減一後馬上檢查是否到達零了
   回答 = MsgBox("時間結束，是否再玩一次", vbYesNo + vbQuestion) '使用msgbox函數呈現說明，並詢問玩者是否再玩一次
   If 回答 = vbYes Then           '比對玩者選的鍵，注意是<vbYes>不可只寫Yes
      剩餘時間 = 5              '如果玩者選<是>的鍵，就把所有狀態恢復原始
      已打次數 = 0               '所以改<剩餘時間><已打次數><打中次數>這幾項
      打中次數 = 0
      打中率 = ""
   Else
      MsgBox "Bye Bye..."        '如果不是選擇<是>，那就是不玩了，顯示再見的訊息
      End                        '把程式結束。
   End If
End If
End Sub
