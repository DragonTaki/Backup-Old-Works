VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "電腦訂位"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10440
   Icon            =   "電腦訂位.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10440
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton 存檔 
      BackColor       =   &H000080FF&
      Caption         =   "存檔"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton 取檔 
      BackColor       =   &H0000FFFF&
      Caption         =   "取檔"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox 檔名 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6180
      TabIndex        =   1
      Text            =   "C:\TTT.TXT"
      Top             =   6600
      Width           =   3495
   End
   Begin VB.CommandButton 座位 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   840
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label 班級座號 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '單線固定
      Caption         =   "表單:Form4"
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
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "存檔檔名:"
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
      Left            =   4140
      TabIndex        =   3
      Top             =   6600
      Width           =   1905
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label 專案說明 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '單線固定
      Caption         =   "電腦訂位程式"
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
      Height          =   480
      Left            =   5088
      TabIndex        =   2
      Top             =   120
      Width           =   2736
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      X1              =   240
      X2              =   10260
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'電腦訂位程式，主要分為 1.為展開物件，2訂位，3存檔，4取檔共四個動作

Private Sub Form_Load()  '程式開始執行時，先做展開物件的動作
For 排 = 1 To 6          '注意<座位>按鈕物件的Index屬要先填0，使其成為陣列
   For 列 = 1 To 8       '預計打開成 6 排，每排 8 列
      Index = (排 - 1) * 8 + (列 - 1) '第一排第一列<Index>為0,第一排第二列<Index>為1，所以依要建的排跟列算出相對的Index值
      If Index > 0 Then   '因為第0個，是本來就有，所以比0還大的都要複製出來
         Load 座位(Index)
      End If
      座位(Index).Left = 座位(0).Left + 座位(0).Width * (排 - 1) '修改新生出來的<Left>值，第一排跟座位(0)同位置，第二排要右移一個座位(0)的寬度，同理第三排要移兩個座位(0)的寬度
      座位(Index).Top = 座位(0).Top - 座位(0).Height * (列 - 1)  '修改新生出來的<Top>值，第一排跟座位(0)同位置，第二排要上移一個座位(0)的高度，同理第三排要移兩個座位(0)的高度，因往上，所以要用減，Y軸越上越小
      座位(Index).Caption = (Index + 1)            '在新生的物件上加上標題，標題就是編號加1，因為電腦習慣從0算起，人習慣最先的一個是1
      座位(Index).BackColor = vbYellow             '先全設為黃色(代表未訂位)，當被點時再改成綠色(代表已訂位)
      座位(Index).Visible = True                   '新生的物件內定看不見，所以最後要把<Visible>改為<True>否則永遠只會出現一個按鈕
   Next                  '每個<For>的相對句就是<Next>
Next            '每個<For>的相對句就是<Next>，有兩個For,所以相對位置也要兩個Next
End Sub

Private Sub 座位_Click(Index As Integer)  '訂位是用 toggle 交互執行的方式
If 座位(Index).BackColor = vbYellow Then   '如果按鈕的被背色是黃色(代表未訂位)，點了要改成訂位
   座位(Index).BackColor = vbGreen         '把點中的改成綠色背景(代表已訂位)因為是陣列物件，所以點任一個都會進本程式，但會把點中那個的Index代入，所以直接修改Index那個的屬性，就是改點中的那一個按鈕
   座位(Index).Caption = (Index + 1) & ":已訂位" '並把該按鈕呈現的字改成<座號>加上":已訂位"的字，注意編號跟字之間要加 & 來連接
Else                                       '如果不是黃色，代表點的時候是綠色(代表已訂位，所以要改成未訂位)
   座位(Index).BackColor = vbYellow        '把背景色改成綠
   座位(Index).Caption = (Index + 1)       '把上面呈現的字換成只有<座號>
End If
End Sub

Private Sub 存檔_Click()   '存檔須注意使用前要開檔，結束時要記得關檔，否則不能執行第二次
Open 檔名.Text For Output As #1  '開檔格式是 Open(開啟) 檔名 For(為了) Output(輸出) 最後要給個編號，一般都是#1，如同時開檔，第二個就編#2
For Index = 0 To 47              '把48個按鈕現況輸出到檔案
   If 座位(Index).BackColor = vbGreen Then '檢查該按鈕的背景色，如果是<綠色>代表已訂位，否則代表<未訂位>
      Print #1, "已訂位"         '循序把各按鈕現況輸出成文字檔，輸出的指令是 Print #1
   Else
      Print #1, "未訂位"         '沒訂位也要輸出，只此輸出的檔就會有48行資料，剛好依序對映到48按鈕現況
   End If
Next
Close #1                         '輸出完，要馬上關檔，並注意<檔名>這物件要給一個有效的檔名，輸完後要到磁機看看輸出結果是否正確
MsgBox "存檔完成", vbInformation '因輸出很快，使用者常不清楚有沒有輸出，所以完成時顯示一個<存檔完成>的訊息框給使用者看
End Sub

Private Sub 取檔_Click()
Open 檔名.Text For Input As #1 '取檔跟存檔很像，只是反向動作，所以<For>後面必須是<Input>
For Index = 0 To 47            '存48個資料，當然就要取回48個資料
   Line Input #1, AAA          '讀回資料指令為 Line Input #1， 先放在一個臨時變數裡
   If AAA = "已訂位" Then      '依檔案內讀入一行的內容，來比較，因為一行是一個座位的資料，
      座位(Index).BackColor = vbGreen                '所以如果字是<已訂位>，代表此座位原來是有訂位，注意此處的字，要和存檔時送出的字完全相同
      座位(Index).Caption = (Index + 1) & ":已訂位"  '把此座位的背景色換成<綠色>，並把標題改成座號加<:已訂位>
   Else
      座位(Index).BackColor = vbYellow       '如果比對不是<已訂位>那就代表未訂位，所以換成黃色背景及只顯示座號
      座位(Index).Caption = (Index + 1)
   End If
Next
Close #1                     '取回完成，立即關檔，使可以執行下次的存檔或取檔
MsgBox "取回完成"            '顯示一句訊息，讓使用者知道資料已取回，注意 Msgbox是指令，所以Msgbox接的不是 <等號>，是<空白>，然後才是要顯示的 "字串"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '當按了表單右上角X時會發生的事件
Form5.Show                                  '因為是多重表單，所以在本表單結束前要啟動選項表單。考試時此處不必寫。
End Sub                                     '考試時此段不必寫

