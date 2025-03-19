VERSION 5.00
Begin VB.Form 規則設定 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "M04_設定.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   13050
   StartUpPosition =   2  '螢幕中央
End
Attribute VB_Name = "規則設定"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp1 As Integer
Dim Temp2 As String
Dim CountTime As Integer
Private Sub Form_Load()
   設定框.Left = 0
   設定框.Top = 0
   Randomize
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   主頁.Show
End Sub
Private Sub 回主頁_Click()
   規則設定.Hide
   主頁.Show
End Sub
Private Sub 炮跳開_Click()
   是否炮跳 = True
End Sub
Private Sub 炮跳關_Click()
   是否炮跳 = False
End Sub
Private Sub 音樂計時器_Timer()
   CountTime = CountTime + 1
   If Temp2 = "0003" Then
      If CountTime = 102 Then
         撥放器.Command = "CLOSE"
         CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "0006" Then
      If CountTime = 73 Then
         撥放器.Command = "CLOSE"
         CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "Credits" Then
      If CountTime = 205 Then
         撥放器.Command = "CLOSE"
         CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "Revolootin" Then
      If CountTime = 72 Then
         撥放器.Command = "CLOSE"
         CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "SomeOfAKind" Then
      If CountTime = 76 Then
         撥放器.Command = "CLOSE"
         CountTime = 0
         Call 音樂播放
      End If
   End If
End Sub
Private Sub 音樂開_Click()
   是否音樂 = True
   Call 音樂播放
End Sub
Private Sub 音樂關_Click()
   是否音樂 = False
   Call 音樂播放
End Sub
Public Sub 音樂播放()
   Temp1 = Int(Rnd * 5)
   Select Case Temp1
      Case 0:
         Temp2 = "0003"
      Case 1:
         Temp2 = "0006"
      Case 2:
         Temp2 = "Credits"
      Case 3:
         Temp2 = "Revolootin"
      Case 4:
         Temp2 = "SomeOfAKind"
   End Select
   If 是否音樂 = True Then
      撥放器.FileName = App.Path & "\" & Temp2 & ".mp3"
      撥放器.Command = "OPEN"
      撥放器.Command = "PLAY"
   Else
      撥放器.Command = "CLOSE"
   End If
End Sub
