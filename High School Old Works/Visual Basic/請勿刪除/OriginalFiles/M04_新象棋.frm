VERSION 5.00
Begin VB.Form �D�� 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H��"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   13050
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton �C���������s 
      BackColor       =   &H0000FFFF&
      Caption         =   "�C������"
      BeginProperty Font 
         Name            =   "�رd�����(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton ������s 
      BackColor       =   &H0000FFFF&
      Caption         =   "���@�@��"
      BeginProperty Font 
         Name            =   "�رd�����(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton �C���]�w 
      BackColor       =   &H0000FFFF&
      Caption         =   "�C���]�w"
      BeginProperty Font 
         Name            =   "�رd�����(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton �i�J�C�� 
      BackColor       =   &H0000FFFF&
      Caption         =   "�i�J�C��"
      BeginProperty Font 
         Name            =   "�رd�����(P)"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Menu �ɮ� 
      Caption         =   "�ɮ�(&F)"
      Begin VB.Menu ���� 
         Caption         =   "����(&X)"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����(&H)"
      Begin VB.Menu �C������ 
         Caption         =   "�C������(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu ���� 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "�D��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Name1P As String
Public Name2P As String
Private Sub Form_Load()
   �O�_�t�Y = False
   �O�_���� = True
   �O�_���� = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   End
End Sub
Private Sub ����_Click()
   MsgBox ("�C�������I")
   End
End Sub
Private Sub �i�J�C��_Click()
   Name1P = InputBox("��J�Ĥ@�쪱�a�m�W", "", "�L�W")
   If Name1P <> "" Then
      Name2P = InputBox("��J�ĤG�쪱�a�m�W", "", "�L�W")
      If Name2P <> "" Then
         �H��.Show
         �D��.Hide
      End If
   End If
End Sub
Private Sub �C���]�w_Click()
   �W�h�]�w.Show
   �D��.Hide
End Sub
Private Sub �C������_Click()
   ������.Show
   �D��.Hide
End Sub
Private Sub �C���������s_Click()
   Call �C������_Click
End Sub
Private Sub ����_Click()
   ����.Show
   �D��.Hide
End Sub
Private Sub ������s_Click()
   Call ����_Click
End Sub
