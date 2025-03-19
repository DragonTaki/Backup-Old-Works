VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   Icon            =   "連結表單.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   8385
   ScaleWidth      =   9420
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton 表單三 
      BackColor       =   &H0080C0FF&
      Caption         =   "貪吃蛇遊戲"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton 表單一 
      BackColor       =   &H0080C0FF&
      Caption         =   "打地鼠遊戲"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton 表單二 
      BackColor       =   &H0080C0FF&
      Caption         =   "打鴨子遊戲"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   2700
      Width           =   2295
   End
   Begin VB.ListBox 題目 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   2010
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   4440
      Width           =   6855
   End
   Begin VB.ListBox 題目 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1230
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   1020
      Width           =   6855
   End
   Begin VB.ListBox 題目 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1815
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   2460
      Width           =   6855
   End
   Begin VB.Label 合計說明 
      AutoSize        =   -1  'True
      Caption         =   "按遊戲名稱可試玩"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   12
      Top             =   7020
      Width           =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "三個遊戲各自獨立寫一個專案表單，彼此無關"
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
      Left            =   300
      TabIndex        =   11
      Top             =   7500
      Width           =   9015
   End
   Begin VB.Label 小計 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   10
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label 小計 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label 小計 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Index           =   2
      Left            =   1620
      TabIndex        =   8
      Top             =   5580
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "按遊戲名稱可試玩"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '單線固定
      Caption         =   "電腦課程VB期未考"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2565
      TabIndex        =   0
      Top             =   480
      Width           =   3825
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 專案說明_Click()

End Sub

Private Sub Form_Load()
i = 0
題目(i).AddItem "10分:地鼠每0.5秒會換一次位置"
題目(i).AddItem " 5分:每打一下，<已打次數>數字會加一"
題目(i).AddItem " 5分:每打中地鼠一下，<打中次數>數字會加一"
題目(i).AddItem " 5分:<打中率>會依每打一下保持正確比率，並有小數兩位的精確度"
題目(i).AddItem " 5分:倒數計時器會每秒倒數一，到零會停止計時"
題目(i).AddItem "10分:倒數停止計時後會問是否再玩一次,並可再玩或結束"
i = i + 1
題目(i).AddItem " 5分:鴨子可以左右自動移動"
題目(i).AddItem " 5分:烏鴉可以左右自動移動"
題目(i).AddItem " 5分:地面飛彈會跟滑鼠在地面左右移"
題目(i).AddItem " 5分:按一下滑鼠，飛彈會起飛往上"
題目(i).AddItem " 5分:每發射一個飛彈<發射數量>的數字會加一"
題目(i).AddItem " 5分:飛出表單上方<狀態>會出現<損失飛彈一枚>並收回飛彈"
題目(i).AddItem "10分:飛彈打中鴨子<狀態>會出現<打中目標>並收回飛彈"
題目(i).AddItem " 5分:每打中鴨子<命中數量>的數字會加一"
題目(i).AddItem " 5分:飛彈打中烏鴉<狀態>會出現<打錯目標>並收回飛彈"
i = i + 1
題目(i).AddItem "10分:按數字2,4,6,8,5<蛇的方向>上的字會依序變成下,左,右,上,停止"
題目(i).AddItem " 5分:執行時蛇及青蛙都能自動變成<蛇(0).Width>的相同正方塊"
題目(i).AddItem " 5分:開始執行時每次蛇及青蛙都可以在表單上不同位置出現"
題目(i).AddItem " 5分:蛇頭會跟著<蛇的方向>走動"
題目(i).AddItem " 5分:蛇會正確的吃到青蛙，青蛙會換位置，蛇會加長一節，速度變快"
題目(i).AddItem "10分:蛇身會正確的跟著蛇頭走動"
題目(i).AddItem " 5分:蛇吃到青蛙，<最高紀錄>會正確更新。"
題目(i).AddItem " 5分:蛇撞到任一面牆會出現<撞到了>且正確收回蛇，可重新再玩一次"
題目(i).AddItem "10分:蛇如果咬到自己會出現<好痛>且正確收回蛇，可重新再玩一次"
滿分 = 0
For j = 0 To 2
   For i = 0 To 題目(j).ListCount - 1
      小計(j) = 小計(j) + Val(題目(j).List(i))
   Next
   滿分 = 滿分 + 小計(j)
Next
合計說明 = "合計 " & 滿分 & " 分，及格為60分，超過90分部份每5分算1分，期末考及格期中考算及格。"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

End

End Sub

Private Sub 表單一_Click()
Form1.Show
Unload Form2
Unload Form3
Form4.Hide

End Sub

Private Sub 表單二_Click()
Unload Form1
Form2.Show
Unload Form3
Form4.Hide

End Sub

Private Sub 表單三_Click()
Unload Form1
Unload Form2
Form3.Show
Form4.Hide

End Sub

