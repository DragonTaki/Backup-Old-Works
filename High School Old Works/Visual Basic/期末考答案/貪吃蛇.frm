VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "貪吃蛇"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   Icon            =   "貪吃蛇.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11220
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton 蛇 
      BackColor       =   &H000080FF&
      Caption         =   "蛇"
      Height          =   1695
      Index           =   0
      Left            =   3120
      Style           =   1  '圖片外觀
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "蛇"
      Top             =   2640
      Width           =   435
   End
   Begin VB.Timer 計時器 
      Interval        =   300
      Left            =   2520
      Top             =   2640
   End
   Begin VB.CommandButton 青蛙 
      BackColor       =   &H0000FF00&
      Caption         =   "蛙"
      Height          =   1605
      Left            =   5820
      Style           =   1  '圖片外觀
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "青蛙"
      Top             =   3120
      Width           =   2205
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '透明
      Caption         =   "最高紀錄:"
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
      Left            =   7800
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label 最高紀錄 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   9
      ToolTipText     =   "蛇的長度"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label 蛇的長度 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "蛇的長度"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label 蛇的方向 
      BackColor       =   &H00FFFF00&
      Caption         =   "停止"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "蛇的方向"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '透明
      Caption         =   "蛇的方向:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label label2 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '透明
      Caption         =   "蛇的長度:"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label 說明標簽 
      BackStyle       =   0  '透明
      Caption         =   "請用2,4,6,8鍵控制方向,5停"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label 專案名稱 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '單線固定
      Caption         =   "貪吃蛇"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label 班級姓名 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '單線固定
      Caption         =   "表單:Form3"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2265
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***注意單引號後面是程式註解說明用，在真正考試寫程式是不必寫的
'記得把表單的<KeyPreview>設為<True>

Dim 格點大小 As Integer   '宣告這個變數，使格點遊戲所有大小及移動距離相同，都用這個值

Private Sub Form_KeyPress(KeyAscii As Integer)   '有人按下按鍵，要使這事件發生，記得把表單的KeyPreview設為True
Select Case Chr(KeyAscii)                       '按鍵按下時會傳入所按的鍵的內碼，就以傳來的內碼當判斷條件
Case "2":                                        '因為傳來是內碼，經CHR()函數換成原入的字
   蛇的方向 = "下"                           '如果相同，代表剛才玩者按了2這個鍵，就把方向設為<下>
Case "4":                                        '同理判斷其他想用的鍵內碼，有相同時
   蛇的方向 = "左"                         '就設成相對的字，以便等下蛇移動時的方向依據。
Case "6":
   蛇的方向 = "右"
Case "8":
   蛇的方向 = "上"
Case "5":
   蛇的方向 = "停"
End Select
End Sub

Private Sub Form_Load()        '程式一開始執行時，會先執行的事件，用來做變數初值設定
Randomize                      '因為有使用到<Rnd>這函數，所以要加這行，使其亂數完全沒有規律性
格點大小 = 蛇(0).Width         '為使所有蛇及青蛙等大，所以先把<蛇(0).Width>存在宣告過的<格點大小>這變數裡
蛇(0).Height = 格點大小        '把蛇頭的高度設成跟寬度一樣，這樣蛇頭就會是正方形
                              
青蛙.Width = 格點大小          '把青蛙也設成等大
青蛙.Height = 格點大小

蛇(0).Left = 取X值             '利用已寫好的格點化函數，來使一開始<蛇(0)>隨意在表單中找一個新位置
蛇(0).Top = 取Y值

青蛙.Left = 取X值              '同樣的，把青蛙也隨機放在表單的某個格點上
青蛙.Top = 取Y值
End Sub

Private Function 取X值()                          '自寫的副程式，用來在表單中隨機找到一個格點大小整倍數的X位置
最大倍數 = Int(Me.Width / 格點大小)             '拿表單寬度除格點大小，算出目前橫的可放幾隻蛇
取X值 = Int(Rnd * 最大倍數) * 格點大小             '依所能放的倍數取Rnd化成整數再乘<格點大小>，這樣保證取到的位置是蛇寬的倍數
End Function
Private Function 取Y值()                           '自寫的副程式，用來在表單中隨機找到一個格點大小整倍數的Y位置
表單實際高度 = Me.Height - 600   '注意是Me.Height不是Width '因為表單有個藍色標題區，會佔用約600的高度，所需減掉,
最大倍數 = Int(表單實際高度 / 格點大小)             '拿表單實際高度，除格點大小，算出目前縱的可放幾隻蛇
取Y值 = Int(Rnd * 最大倍數) * 格點大小             '不過由於表單的Y軸是由藍色標題下方算起，所以要先減一個值，以免剛好放在看不到的地方
End Function                                       '依所能放的倍數取Rnd化成整數再乘<格點大小>，這樣保證取到的位置是蛇寬的倍數

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '當按了表單右上角X時會發生的事件
Form4.Show                                  '因為是多重表單，所以在本表單結束前要啟動選項表單。考試時此處不必寫。
End Sub                                     '考試時此段不必寫

Private Sub 計時器_Timer()          '固定時間週期內會來執行一遍的事件
If Me.WindowState = vbMinimized Then
   Exit Sub
End If
'下面是先把蛇頭位置記住，以便移蛇身程式使用
蛇頭X = 蛇(0).Left                      '移動蛇頭前先把目前位標記在X,Y變數裡，以便後面的蛇身能跟在蛇頭後方
蛇頭Y = 蛇(0).Top

'下面是處理<蛇頭移動>的程式
Select Case 蛇的方向                    '依據<蛇的方向>目前呈現的字來決定蛇要往那個方向走一個<格點大小>距離
Case "上":
   蛇(0).Top = 蛇(0).Top - 格點大小     '如果是上下，則改變的是蛇頭的上界<蛇(0).Top>，每次改變一個蛇身距離
Case "下":
   蛇(0).Top = 蛇(0).Top + 格點大小
Case "左":
   蛇(0).Left = 蛇(0).Left - 格點大小   '如果是左右，則改變的是蛇頭的上界<蛇(0).Left>，每次改變一個蛇身距離
Case "右":
   蛇(0).Left = 蛇(0).Left + 格點大小   '停不必處理，因為如果是停就不必移動。
End Select

'下面是處理<蛇吃到青蛙>的程式                             '蛇頭移動完，蛇身還沒動時先判斷有沒有吃到青蛙，以便先把蛇身加長
If 蛇(0).Top = 青蛙.Top And 蛇(0).Left = 青蛙.Left Then   '如果青蛙與蛇(0)的上界跟左界相同代表蛇吃到青蛙了
   青蛙.Left = 取X值                                      '青蛙被咬到，就替青蛙換個新的位置
   青蛙.Top = 取Y值
   Load 蛇(蛇的長度)                                      '複製一個蛇身
   蛇(蛇的長度).Caption = "X"                              '使蛇身上沒有字
   蛇(蛇的長度).BackColor = vbYellow                      '使蛇身跟蛇頭不同色
   蛇(蛇的長度).Visible = True                            '使新產生的這個蛇身出現
   計時器.Interval = 計時器.Interval * 0.8                '為了增加效果，吃一隻青蛙就讓蛇加速百分之20
   蛇的長度 = Val(蛇的長度) + 1                           '並把<蛇的長度>加一，以便下次吃到青蛙時會呈現下一節
   If Val(最高紀錄) < Val(蛇的長度) Then                  '<最高紀錄>與<>都是字串，所以一定要先用val化成數值來比較
      最高紀錄 = 蛇的長度                                 '否則會發生<"9">比<"10">大的現象。
   End If                                                 '可是要改變某個變數時，等號左邊的變數卻不可加上val,只直接用變數名稱
End If


'下面是處理<蛇身移動>的程式
For i = 1 To 蛇的長度 - 1                                 '移動蛇身程式，讓每一節都跟著前一節走,因為只要處蛇身所以 For 是從 1 算起的。
   蛇身X = 蛇(i).Left                                     '要往前一位置移之前先把目前座標留下在X2,Y2
   蛇身Y = 蛇(i).Top
   蛇(i).Left = 蛇頭X                                     '真的把目前這蛇移到蛇頭位置
   蛇(i).Top = 蛇頭Y
   蛇頭X = 蛇身X                                          '硬把X,Y位置設為剛移走這節蛇的位置，使下一節蛇身誤以為
   蛇頭Y = 蛇身Y                                          '他前面的那節就是蛇頭，這樣就會一節跟著一節走
Next

'下面是處理<蛇撞牆>程式
If 蛇(0).Left < 0 Or 蛇(0).Left > Me.Width Or 蛇(0).Top < 0 Or 蛇(0).Top > Me.Height - 600 Then '判斷是否撞牆
   MsgBox "撞到了"                       '如果蛇(0)的左界跑到表單的左右外面或蛇(0)的上界跑到表單上下的外面
   蛇(0).Left = 取X值                    '代表撞到了，這時有六件事要做，以恢復蛇的起始狀態。
   蛇(0).Top = 取Y值                     '一.秀個訊息 用msgbox指令，
   蛇的方向 = "停"               '二.替蛇(0)找一個新的位
   For i = 1 To 蛇的長度 - 1             '三.因為有些蛇身已呈現，再把所有蛇身全部拿掉
      Unload 蛇(i)                       '四.把蛇的長度改成基本長度，就是1節
   Next
   蛇的長度 = 1                  '五.因為重定位，怕玩者還沒看清楚，就撞牆了，所以<蛇的方向>設為"停"。
   計時器.Interval = 300                 '六.恢復原始移動速度，以免重玩，一開始就很快
End If
'下面是處理<咬蛇自盡>程式
If 蛇的方向 <> "停" Then                  '判斷是否蛇有咬到自己，不過首先必須不在<停>的狀態中，否則按<停>就咬到了
   有咬到自己 = False
   For i = 1 To 蛇的長度 - 1             '實際有蛇身長度蛇頭不算，所以只判斷到<蛇的長度>減一節
      If 蛇(0).Left = 蛇(i).Left And 蛇(0).Top = 蛇(i).Top Then  '如果<蛇(0)>跟蛇的某節<蛇(J)>的左界跟上界相同就代表咬到自己了
         有咬到自己 = True
      End If
   Next
   If 有咬到自己 = True Then
      MsgBox "好痛"                   '咬到自己跟前面撞到牆一樣要做六件事來重新啟動貪吃蛇
      蛇(0).Left = 取X值
      蛇(0).Top = 取Y值
      蛇的方向 = "停"
      For i = 1 To 蛇的長度 - 1
         Unload 蛇(i)
      Next
      蛇的長度 = 1
      計時器.Interval = 300           '因為是一節節檢查蛇身是否有跟蛇頭在同位置，有咬到自己時會把蛇身消除
   End If
End If
End Sub
