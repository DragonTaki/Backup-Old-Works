VERSION 5.00
Begin VB.Form �}�l��� 
   BorderStyle     =   0  '�S���ؽu
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12945
   Icon            =   "M04_�}�l���.frx":0000
   Picture         =   "M04_�}�l���.frx":324A
   ScaleHeight     =   9660
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Timer �}�l�p�ɾ� 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "�}�l���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ���w�� As Boolean
Dim CountTime As Integer
Dim MakeFile As Boolean
Public �ɮת��� As String
Public �ɮת����ƭ� As Double '���p��
Private Sub Form_Load()
   Dim TheFile As String
   Dim Results As String
   MakeFile = 1 'False/0:MakeExe True/1:MakeInstallion
   �ɮת��� = "V2.3" '�C���Ч�惡''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   �ɮת����ƭ� = 2.3 '�C���Ч�惡''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   TheFile = "C:\WINDOWS\ChineseChess.ini"
   Results = Dir$(TheFile)
   If Results = "" Then '�ɮפ��s�b
      ���w�� = True
      �w�˭�.Show
      Unload �}�l���
   Else '�ɮצs�b
      Open "C:\WINDOWS\ChineseChess.ini" For Input As #1 'Ū��
         Line Input #1, Temp
         If Val(Temp) <> Val(�ɮת����ƭ�) Then '�D������
            Close #1
            �w�˭�.Show
            Unload �}�l���
         Else
            If MakeFile = False Then
               Call MakeExe
            Else 'MakeFile = True
               Call MakeInstallion
            End If
         End If
      Close #1
   End If
   'SetAttr "C:\WINDOWS\ChineseChess.txt", vbHidden '����
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   CountTime = 1
End Sub
Private Sub �}�l�p�ɾ�_Timer()
   CountTime = CountTime + 1
   If CountTime >= 1 Then
      �D��.Show
      Unload �}�l���
   End If
End Sub
Private Sub MakeExe()
   �}�l�p�ɾ�.Enabled = True '�}�l
End Sub
Private Sub MakeInstallion()
   Close #1 '�ק�״_�����{��
   �w�˭�.Show
   Unload �}�l���
End Sub
