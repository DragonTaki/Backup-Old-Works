VERSION 5.00
Begin VB.Form �W�h�]�w 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H��"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "M04_�]�w.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   13050
   StartUpPosition =   2  '�ù�����
End
Attribute VB_Name = "�W�h�]�w"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp1 As Integer
Dim Temp2 As String
Dim CountTime As Integer
Private Sub Form_Load()
   �]�w��.Left = 0
   �]�w��.Top = 0
   Randomize
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   �D��.Show
End Sub
Private Sub �^�D��_Click()
   �W�h�]�w.Hide
   �D��.Show
End Sub
Private Sub �����}_Click()
   �O�_���� = True
End Sub
Private Sub ������_Click()
   �O�_���� = False
End Sub
Private Sub ���֭p�ɾ�_Timer()
   CountTime = CountTime + 1
   If Temp2 = "0003" Then
      If CountTime = 102 Then
         ����.Command = "CLOSE"
         CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "0006" Then
      If CountTime = 73 Then
         ����.Command = "CLOSE"
         CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "Credits" Then
      If CountTime = 205 Then
         ����.Command = "CLOSE"
         CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "Revolootin" Then
      If CountTime = 72 Then
         ����.Command = "CLOSE"
         CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "SomeOfAKind" Then
      If CountTime = 76 Then
         ����.Command = "CLOSE"
         CountTime = 0
         Call ���ּ���
      End If
   End If
End Sub
Private Sub ���ֶ}_Click()
   �O�_���� = True
   Call ���ּ���
End Sub
Private Sub ������_Click()
   �O�_���� = False
   Call ���ּ���
End Sub
Public Sub ���ּ���()
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
   If �O�_���� = True Then
      ����.FileName = App.Path & "\" & Temp2 & ".mp3"
      ����.Command = "OPEN"
      ����.Command = "PLAY"
   Else
      ����.Command = "CLOSE"
   End If
End Sub
