VERSION 5.00
Begin VB.Form ������ 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H��"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "M04_����.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13050
   StartUpPosition =   2  '�ù�����
End
Attribute VB_Name = "������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   �D��.Show
End Sub
Private Sub �����T�w_Click()
   �D��.Show
   ������.Hide
End Sub
