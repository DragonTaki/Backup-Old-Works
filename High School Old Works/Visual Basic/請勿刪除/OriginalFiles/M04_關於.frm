VERSION 5.00
Begin VB.Form 關於頁 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "M04_關於.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13050
   StartUpPosition =   2  '螢幕中央
End
Attribute VB_Name = "關於頁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   關於頁框.Left = 0
   關於頁框.Top = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   主頁.Show
End Sub
Private Sub 連結_Click()
   Shell "C:\Program Files\Internet Explorer\iexplore.exe & http://www.facebook.com/#!/groups/269973856388003/", vbMaximizedFocus
End Sub
Private Sub 連結_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   連結.ForeColor = vbRed
End Sub
Private Sub 連結_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   連結.ForeColor = &H80FF&
End Sub
Private Sub 確定_Click()
   主頁.Show
   關於頁.Hide
End Sub
