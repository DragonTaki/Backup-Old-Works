VERSION 5.00
Begin VB.Form �H�� 
   BackColor       =   &H005FA76F&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H��"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13050
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame ���� 
      BackColor       =   &H005FA76F&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�رd������W12"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1935
      Left            =   4680
      TabIndex        =   5
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Frame �Ѿl�Ѽ� 
      BackColor       =   &H005FA76F&
      Caption         =   "�Ѿl�Ѽ�"
      BeginProperty Font 
         Name            =   "�رd������W12"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1935
      Left            =   480
      TabIndex        =   4
      Top             =   7560
      Width           =   3855
   End
   Begin VB.Image �����a�G�� 
      Height          =   570
      Left            =   7129
      Picture         =   "M04_�H��0.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image �����a�@�� 
      Height          =   570
      Left            =   1129
      Picture         =   "M04_�H��0.frx":0F5B
      Stretch         =   -1  'True
      Top             =   840
      Width           =   555
   End
   Begin VB.Label ���a�G�m�W 
      Alignment       =   2  '�m�����
      BackColor       =   &H005FA76F&
      BeginProperty Font 
         Name            =   "�رd������(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   9747
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label ���a�@�m�W 
      Alignment       =   2  '�m�����
      BackColor       =   &H005FA76F&
      BeginProperty Font 
         Name            =   "�رd������(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3754
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label ���a�G 
      Alignment       =   2  '�m�����
      BackColor       =   &H005FA76F&
      Caption         =   "���a�G"
      BeginProperty Font 
         Name            =   "�رd�ʶ���(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7707
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label ���a�@ 
      Alignment       =   2  '�m�����
      BackColor       =   &H005FA76F&
      Caption         =   "���a�@"
      BeginProperty Font 
         Name            =   "�رd�ʶ���(P)"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1714
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   11805
      X2              =   11805
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   10485
      X2              =   10485
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   9165
      X2              =   9165
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   7845
      X2              =   7845
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   6525
      X2              =   6525
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   5205
      X2              =   5205
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   3885
      X2              =   3885
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   2565
      X2              =   2565
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   1245
      Y1              =   1920
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   1245
      X2              =   11805
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   0
      Left            =   1365
      Picture         =   "M04_�H��0.frx":1EB6
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   1
      Left            =   2685
      Picture         =   "M04_�H��0.frx":2019
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   16
      Left            =   1365
      Picture         =   "M04_�H��0.frx":217C
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   8
      Left            =   1365
      Picture         =   "M04_�H��0.frx":22DF
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   9
      Left            =   2685
      Picture         =   "M04_�H��0.frx":2442
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   2
      Left            =   4005
      Picture         =   "M04_�H��0.frx":25A5
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   3
      Left            =   5325
      Picture         =   "M04_�H��0.frx":2708
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   4
      Left            =   6645
      Picture         =   "M04_�H��0.frx":286B
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   5
      Left            =   7965
      Picture         =   "M04_�H��0.frx":29CE
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   6
      Left            =   9285
      Picture         =   "M04_�H��0.frx":2B31
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   7
      Left            =   10605
      Picture         =   "M04_�H��0.frx":2C94
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   15
      Left            =   10605
      Picture         =   "M04_�H��0.frx":2DF7
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   14
      Left            =   9285
      Picture         =   "M04_�H��0.frx":2F5A
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   13
      Left            =   7965
      Picture         =   "M04_�H��0.frx":30BD
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   12
      Left            =   6645
      Picture         =   "M04_�H��0.frx":3220
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   11
      Left            =   5325
      Picture         =   "M04_�H��0.frx":3383
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   10
      Left            =   4005
      Picture         =   "M04_�H��0.frx":34E6
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   23
      Left            =   10605
      Picture         =   "M04_�H��0.frx":3649
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   22
      Left            =   9285
      Picture         =   "M04_�H��0.frx":37AC
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   21
      Left            =   7965
      Picture         =   "M04_�H��0.frx":390F
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   20
      Left            =   6645
      Picture         =   "M04_�H��0.frx":3A72
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   19
      Left            =   5325
      Picture         =   "M04_�H��0.frx":3BD5
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   18
      Left            =   4005
      Picture         =   "M04_�H��0.frx":3D38
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   17
      Left            =   2685
      Picture         =   "M04_�H��0.frx":3E9B
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   31
      Left            =   10605
      Picture         =   "M04_�H��0.frx":3FFE
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   30
      Left            =   9285
      Picture         =   "M04_�H��0.frx":4161
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   29
      Left            =   7965
      Picture         =   "M04_�H��0.frx":42C4
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   28
      Left            =   6645
      Picture         =   "M04_�H��0.frx":4427
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   27
      Left            =   5325
      Picture         =   "M04_�H��0.frx":458A
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   26
      Left            =   4005
      Picture         =   "M04_�H��0.frx":46ED
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   25
      Left            =   2685
      Picture         =   "M04_�H��0.frx":4850
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Image �� 
      Height          =   1215
      Index           =   24
      Left            =   1365
      Picture         =   "M04_�H��0.frx":49B3
      Stretch         =   -1  'True
      Tag             =   "-1"
      Top             =   6000
      Width           =   1080
   End
End
Attribute VB_Name = "�H��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim �w���Ĥ@���� As Boolean
Dim �w�}�� As Boolean '���C���
Dim �Ѥw½�}(32) As Boolean '�Φ�l�O
Dim �Ů�(32) As Boolean '�Φ�l�O
Dim �{�b�����a�G As Boolean
Dim ���a�@�O�ª� As Boolean
Dim �Ĥ@���O�ۤv���� As Boolean
Dim �ĤG���O�ۤv���� As Boolean
Dim �O�i���ʪ� As Boolean
Dim �Ĥ@���� As Integer
Dim �ĤG���� As Integer
Dim �Ĥ@���Ѧ�m As Integer
Dim �ĤG���Ѧ�m As Integer
Dim �������� As Integer
Private Sub Form_Load()
   ���a�@�m�W.Caption = �D��.Name1P
   ���a�G�m�W.Caption = �D��.Name2P
   Randomize
   Call �~�P
End Sub
Private Sub �~�P()
   Dim �Ѥw�o�L(32) As Boolean
   For i = 0 To 31
      �Ѥw�o�L(i) = False
   Next
   For i = 0 To 31
�����:
      �n�񪺴� = Int(Rnd * 32)
      If �Ѥw�o�L(�n�񪺴�) = True Then
         GoTo �����
      End If
      �Ѥw�o�L(�n�񪺴�) = True
      ��(i).Tag = �n�񪺴�
   Next
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Ans = MsgBox("�A�T�w�n�����H", vbYesNo, "")
   If Ans = vbYes Then
      �D��.Show
      �H��.Hide
      Name1P = Empty
      Name2P = Empty
   Else
      Cancel = 1
   End If
End Sub
Private Sub ��_Click(Index As Integer)
   If �w�}�� = False Then '�M�w�C��
      If ��(Index).Tag = 0 Or ��(Index).Tag = 1 Or ��(Index).Tag = 2 Or ��(Index).Tag = 3 Or ��(Index).Tag = 4 Or ��(Index).Tag = 5 Or ��(Index).Tag = 6 Or ��(Index).Tag = 7 Or ��(Index).Tag = 8 Or ��(Index).Tag = 9 Or ��(Index).Tag = 10 Or ��(Index).Tag = 11 Or ��(Index).Tag = 12 Or ��(Index).Tag = 13 Or ��(Index).Tag = 14 Or ��(Index).Tag = 15 Then '�I��´�
         ���a�@.ForeColor = vbBlack
         ���a�G.ForeColor = vbRed
         ���a�@�O�ª� = True
      Else '�I�����
         ���a�@.ForeColor = vbRed
         ���a�G.ForeColor = vbBlack
      End If
      Call ���H
      �w�}�� = True
      ��(Index).Picture = LoadPicture(App.Path & "\" & ��(Index).Tag & ".gif")
      �Ѥw½�}(Index) = True
   Else '�w�M�w�C��
      If �Ѥw½�}(Index) = False Then '�Ӵѥ�½�}
         ��(Index).Picture = LoadPicture(App.Path & "\" & ��(Index).Tag & ".gif")
         �Ѥw½�}(Index) = True
         Call ���H
      ElseIf �w���Ĥ@���� = False Then '�ӴѤw½�}�A�O�Ĥ@��
         �w���Ĥ@���� = True
         �Ĥ@���� = ��(Index).Tag
         �Ĥ@���Ѧ�m = Index
         Call �P�_�Ĥ@���O�_���ۤv����
         If �Ĥ@���O�ۤv���� = False Then '���O�ۤv����
            �w���Ĥ@���� = False '�٨S��
            �Ĥ@���� = Empty
            �Ĥ@���Ѧ�m = Empty
         Else '�O�ۤv����
            �Ĥ@���O�ۤv���� = False
         End If
      ElseIf �w���Ĥ@���� = True And �Ѥw½�}(Index) = False Then '�ĤG���ѥ�½�}
         If �O�_�t�Y = True Then '�i�t�Y
            GoTo �i�t�Y '���P�w½�}
         End If
         '�٨S��
      ElseIf �w���Ĥ@���� = True And �Ѥw½�}(Index) = True Then '�ӴѤw½�}�A�O�ĤG��
�i�t�Y:
         �ĤG���� = ��(Index).Tag
         �ĤG���Ѧ�m = Index
         Call �P�_�ĤG���O�_���ۤv����
         If �ĤG���O�ۤv���� = True Then
            �ĤG���O�ۤv���� = False
            �ĤG���� = Empty
            �ĤG���Ѧ�m = Empty
            Exit Sub '�ĤG���O�ۤv���ѡA����Y�A���}
         Else
            '�ĤG�����O�ۤv���ѡA�i�H�Y
         End If
         Call �ˬd�����Ѫ���m
         If �O�i���ʪ� = False Then
            Exit Sub
         End If
         �O�i���ʪ� = False
         �������� = 0
         Call �ʴѤl
         �w���Ĥ@���� = False
         �Ĥ@���� = Empty
         �ĤG���� = Empty
         �Ĥ@���Ѧ�m = Empty
         �ĤG���Ѧ�m = Empty
         Call ���H
      End If
   End If
End Sub
Private Sub �P�_�Ĥ@���O�_���ۤv����()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If �Ĥ@���� = 0 Or �Ĥ@���� = 1 Or �Ĥ@���� = 2 Or �Ĥ@���� = 3 Or �Ĥ@���� = 4 Or �Ĥ@���� = 5 Or �Ĥ@���� = 6 Or �Ĥ@���� = 7 Or �Ĥ@���� = 8 Or �Ĥ@���� = 9 Or �Ĥ@���� = 10 Or �Ĥ@���� = 11 Or �Ĥ@���� = 12 Or �Ĥ@���� = 13 Or �Ĥ@���� = 14 Or �Ĥ@���� = 15 Then '�Ĥ@���O�´�
            �Ĥ@���O�ۤv���� = True
         End If
      Else '���a�@�O����
         If �Ĥ@���� = 16 Or �Ĥ@���� = 17 Or �Ĥ@���� = 18 Or �Ĥ@���� = 19 Or �Ĥ@���� = 20 Or �Ĥ@���� = 21 Or �Ĥ@���� = 22 Or �Ĥ@���� = 23 Or �Ĥ@���� = 24 Or �Ĥ@���� = 25 Or �Ĥ@���� = 26 Or �Ĥ@���� = 27 Or �Ĥ@���� = 28 Or �Ĥ@���� = 29 Or �Ĥ@���� = 30 Or �Ĥ@���� = 31 Then '�Ĥ@���O����
            �Ĥ@���O�ۤv���� = True
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If �Ĥ@���� = 16 Or �Ĥ@���� = 17 Or �Ĥ@���� = 18 Or �Ĥ@���� = 19 Or �Ĥ@���� = 20 Or �Ĥ@���� = 21 Or �Ĥ@���� = 22 Or �Ĥ@���� = 23 Or �Ĥ@���� = 24 Or �Ĥ@���� = 25 Or �Ĥ@���� = 26 Or �Ĥ@���� = 27 Or �Ĥ@���� = 28 Or �Ĥ@���� = 29 Or �Ĥ@���� = 30 Or �Ĥ@���� = 31 Then '�Ĥ@���O����
            �Ĥ@���O�ۤv���� = True
         End If
      Else '���a�G�O�ª�
         If �Ĥ@���� = 0 Or �Ĥ@���� = 1 Or �Ĥ@���� = 2 Or �Ĥ@���� = 3 Or �Ĥ@���� = 4 Or �Ĥ@���� = 5 Or �Ĥ@���� = 6 Or �Ĥ@���� = 7 Or �Ĥ@���� = 8 Or �Ĥ@���� = 9 Or �Ĥ@���� = 10 Or �Ĥ@���� = 11 Or �Ĥ@���� = 12 Or �Ĥ@���� = 13 Or �Ĥ@���� = 14 Or �Ĥ@���� = 15 Then '�Ĥ@���O�´�
            �Ĥ@���O�ۤv���� = True
         End If
      End If
   End If
End Sub
Private Sub �P�_�ĤG���O�_���ۤv����()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Ĥ@���O�´�
            �ĤG���O�ۤv���� = True
         End If
      Else '���a�@�O����
         If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then '�Ĥ@���O����
            �ĤG���O�ۤv���� = True
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If �Ĥ@���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then '�Ĥ@���O����
            �ĤG���O�ۤv���� = True
         End If
      Else '���a�G�O�ª�
         If �Ĥ@���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Ĥ@���O�´�
            �ĤG���O�ۤv���� = True
         End If
      End If
   End If
End Sub
Private Sub ���H()
      If �{�b�����a�G = False Then
         �{�b�����a�G = True
         �����a�@��.Visible = False
         �����a�G��.Visible = True
      Else
         �{�b�����a�G = False
         �����a�@��.Visible = True
         �����a�G��.Visible = False
      End If
End Sub
Private Sub �ˬd�����Ѫ���m()
   If �O�_���� = False Then '������
      GoTo ���}���� '�M�@��Ѱʪk�@��
   End If
   If �Ĥ@���� <> 9 And �Ĥ@���� <> 10 And �Ĥ@���� <> 25 And �Ĥ@���� <> 26 Then '�Ĥ@���Ѥ��O�]�ά�
���}����:
      If �Ĥ@���Ѧ�m Mod 8 = �ĤG���Ѧ�m Mod 8 And �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 8 Then '�ĤG���Ѧb�Ĥ@���ѤW��
         �O�i���ʪ� = True
      ElseIf �Ĥ@���Ѧ�m Mod 8 = �ĤG���Ѧ�m Mod 8 And �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 8 Then '�ĤG���Ѧb�Ĥ@���ѤU��
         �O�i���ʪ� = True
      ElseIf �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 1 Then '�ĤG���Ѧb�Ĥ@���ѥk��
         �O�i���ʪ� = True
      ElseIf �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 1 Then '�ĤG���Ѧb�Ĥ@���ѥ���
         �O�i���ʪ� = True
      End If
   Else '�Ĥ@���ѬO�]�ά�
      If Abs((�Ĥ@���Ѧ�m - �ĤG���Ѧ�m)) < 8 Then '�Ĥ@���ѩM�ĤG���ѦP��
         If ��(�ĤG���Ѧ�m).Tag = -1 Then '�ĤG���ѬO�Ū�
            If �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 8 Or �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 8 Then '�ĤG���ѬO�Ū��B�b�Ĥ@���ѤW��ΤU��
               �O�i���ʪ� = True
            End If
         Else '�ĤG���Ѥ��O�Ū�
            temp1 = �Ĥ@���Ѧ�m
            temp2 = �ĤG���Ѧ�m
            If temp1 > temp2 Then
               temp1 = �ĤG���Ѧ�m
               temp2 = �Ĥ@���Ѧ�m
            End If
            For i = temp1 + 1 To temp2
               If ��(�ĤG���Ѧ�m).Tag <> -1 Then
                  �������� = �������� + 1
               End If
            Next
            If �������� = 2 Then
               �O�i���ʪ� = True
            End If
         End If
      ElseIf (�Ĥ@���Ѧ�m - �ĤG���Ѧ�m) Mod 8 = 0 Then  '�Ĥ@���ѩM�ĤG���ѦP�C
         If ��(�ĤG���Ѧ�m).Tag = -1 Then '�ĤG���ѬO�Ū�
            If �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 1 Or �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 1 Then '�ĤG���ѬO�Ū��B�b�Ĥ@���ѥ���Υk��
               �O�i���ʪ� = True
            End If
         Else '�ĤG���Ѥ��O�Ū�
            temp1 = �Ĥ@���Ѧ�m
            temp2 = �ĤG���Ѧ�m
            If temp1 > temp2 Then
               temp1 = �ĤG���Ѧ�m
               temp2 = �Ĥ@���Ѧ�m
            End If
            For i = temp1 + 1 To temp2 Step 8
               If ��(�ĤG���Ѧ�m).Tag <> -1 Then
                  �������� = �������� + 1
               End If
            Next
            If �������� = 2 Then
               �O�i���ʪ� = True
            End If
         End If
      End If
   End If
End Sub
Private Sub �ʴѤl()
   If �{�b�����a�G = False Then
      If ���a�@�O�ª� = True Then
         Call �ʶ´�
      Else '���a�@�O����
         Call �ʬ���
      End If
   Else    '�{�b�����a�G
      If ���a�@�O�ª� = True Then
         Call �ʬ���
      Else '���a�G�O�ª�
         Call �ʶ´�
      End If
   End If
End Sub
Private Sub �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag()
   ��(�ĤG���Ѧ�m).Tag = ��(�Ĥ@���Ѧ�m).Tag
   ��(�Ĥ@���Ѧ�m).Tag = -1
End Sub
Private Sub �ʶ´�()
   If �Ĥ@���� = 0 Then '�N
      If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then '�N�Y�L
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 1 Or �Ĥ@���� = 2 Then '�h
      If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 16 Then  '�h�Y��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 3 Or �Ĥ@���� = 4 Then '�H
      If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Then  '�H�Y�ӥK
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 5 Or �Ĥ@���� = 6 Then '��
      If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then   '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Then  '���Y�ӥK��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 7 Or �Ĥ@���� = 8 Then '��
      If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then   '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Then  '���Y�ӥK����
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 9 Or �Ĥ@���� = 10 Then '�]
      If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y�P��
         ''
      ElseIf �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then   '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 11 Or �Ĥ@���� = 12 Or �Ĥ@���� = 13 Or �Ĥ@���� = 14 Or �Ĥ@���� = 15 Then '��
      If �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Or �ĤG���� = 25 Or �ĤG���� = 26 Then '�Y�P��ά�
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 16 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then   '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Then    '��Y�K���ϽX
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
End Sub
Private Sub �ʬ���()
   If �Ĥ@���� = 16 Then '��
      If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�ӦY��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 17 Or �Ĥ@���� = 18 Then '�K
      If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 0 Then '�K�Y�N
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 19 Or �Ĥ@���� = 20 Then '��
      If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Then '�ۦY�N�h
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 21 Or �Ĥ@���� = 22 Then '��
      If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Then '�ۦY�N�h�H
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 23 Or �Ĥ@���� = 24 Then '�X
      If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y�P��
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Then '�ۦY�N�h�H��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 25 Or �Ĥ@���� = 26 Then '��
      If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then  '�Y�P��
         ''
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
   If �Ĥ@���� = 27 Or �Ĥ@���� = 28 Or �Ĥ@���� = 29 Or �Ĥ@���� = 30 Or �Ĥ@���� = 31 Then '�L
      If �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Or �ĤG���� = 19 Or �ĤG���� = 20 Or �ĤG���� = 21 Or �ĤG���� = 22 Or �ĤG���� = 23 Or �ĤG���� = 24 Or �ĤG���� = 25 Or �ĤG���� = 26 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Or �ĤG���� = 9 Or �ĤG���� = 10 Then  '�Y�P��Υ]
         ''
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Or �ĤG���� = 0 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".gif")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      ElseIf �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Then  '�L�Y�h�H����
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
      End If
   End If
End Sub

