VERSION 5.00
Begin VB.Form 主頁 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   13050
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton 遊戲說明按鈕 
      BackColor       =   &H0000FFFF&
      Caption         =   "遊戲說明"
      BeginProperty Font 
         Name            =   "華康行書體(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton 關於按鈕 
      BackColor       =   &H0000FFFF&
      Caption         =   "關　　於"
      BeginProperty Font 
         Name            =   "華康行書體(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton 遊戲設定 
      BackColor       =   &H0000FFFF&
      Caption         =   "遊戲設定"
      BeginProperty Font 
         Name            =   "華康行書體(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton 進入遊戲 
      BackColor       =   &H0000FFFF&
      Caption         =   "進入遊戲"
      BeginProperty Font 
         Name            =   "華康行書體(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Menu 檔案 
      Caption         =   "檔案(&F)"
      Begin VB.Menu 結束 
         Caption         =   "結束(&X)"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu 說明 
      Caption         =   "說明(&H)"
      Begin VB.Menu 遊戲說明 
         Caption         =   "遊戲說明(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu 關於 
         Caption         =   "關於(&A)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "主頁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Name1P As String
Public Name2P As String
Private Sub Form_Load()
   是否暗吃 = False
   是否炮跳 = True
   是否音樂 = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   End
End Sub
Private Sub 結束_Click()
   MsgBox ("遊戲結束！")
   End
End Sub
Private Sub 進入遊戲_Click()
   Name1P = InputBox("輸入第一位玩家姓名", "", "無名")
   If Name1P <> "" Then
      Name2P = InputBox("輸入第二位玩家姓名", "", "無名")
      If Name2P <> "" Then
         象棋.Show
         主頁.Hide
      End If
   End If
End Sub
Private Sub 遊戲設定_Click()
   規則設定.Show
   主頁.Hide
End Sub
Private Sub 遊戲說明_Click()
   說明頁.Show
   主頁.Hide
End Sub
Private Sub 遊戲說明按鈕_Click()
   Call 遊戲說明_Click
End Sub
Private Sub 關於_Click()
   關於頁.Show
   主頁.Hide
End Sub
Private Sub 關於按鈕_Click()
   Call 關於_Click
End Sub
