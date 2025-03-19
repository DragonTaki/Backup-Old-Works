VERSION 5.00
Begin VB.Form 說明頁 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "M04_說明.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13050
   StartUpPosition =   2  '螢幕中央
End
Attribute VB_Name = "說明頁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   主頁.Show
End Sub
Private Sub 說明確定_Click()
   主頁.Show
   說明頁.Hide
End Sub
