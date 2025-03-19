VERSION 5.00
Begin VB.Form 開始表單 
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12945
   Icon            =   "M04_開始表單.frx":0000
   Picture         =   "M04_開始表單.frx":324A
   ScaleHeight     =   9660
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer 開始計時器 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "開始表單"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 須安裝 As Boolean
Dim CountTime As Integer
Dim MakeFile As Boolean
Public 檔案版本 As String
Public 檔案版本數值 As Double '有小數
Private Sub Form_Load()
   Dim TheFile As String
   Dim Results As String
   MakeFile = 1 'False/0:MakeExe True/1:MakeInstallion
   檔案版本 = "V2.3" '每次請更改此''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   檔案版本數值 = 2.3 '每次請更改此''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   TheFile = "C:\WINDOWS\ChineseChess.ini"
   Results = Dir$(TheFile)
   If Results = "" Then '檔案不存在
      須安裝 = True
      安裝頁.Show
      Unload 開始表單
   Else '檔案存在
      Open "C:\WINDOWS\ChineseChess.ini" For Input As #1 '讀檔
         Line Input #1, Temp
         If Val(Temp) <> Val(檔案版本數值) Then '非此版本
            Close #1
            安裝頁.Show
            Unload 開始表單
         Else
            If MakeFile = False Then
               Call MakeExe
            Else 'MakeFile = True
               Call MakeInstallion
            End If
         End If
      Close #1
   End If
   'SetAttr "C:\WINDOWS\ChineseChess.txt", vbHidden '隱藏
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   CountTime = 1
End Sub
Private Sub 開始計時器_Timer()
   CountTime = CountTime + 1
   If CountTime >= 1 Then
      主頁.Show
      Unload 開始表單
   End If
End Sub
Private Sub MakeExe()
   開始計時器.Enabled = True '開始
End Sub
Private Sub MakeInstallion()
   Close #1 '修改修復移除程式
   安裝頁.Show
   Unload 開始表單
End Sub
