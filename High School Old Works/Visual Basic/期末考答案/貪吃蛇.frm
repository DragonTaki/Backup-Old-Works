VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "�g�Y�D"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   Icon            =   "�g�Y�D.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11220
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton �D 
      BackColor       =   &H000080FF&
      Caption         =   "�D"
      Height          =   1695
      Index           =   0
      Left            =   3120
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "�D"
      Top             =   2640
      Width           =   435
   End
   Begin VB.Timer �p�ɾ� 
      Interval        =   300
      Left            =   2520
      Top             =   2640
   End
   Begin VB.CommandButton �C�� 
      BackColor       =   &H0000FF00&
      Caption         =   "��"
      Height          =   1605
      Left            =   5820
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "�C��"
      Top             =   3120
      Width           =   2205
   End
   Begin VB.Label Label4 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '�z��
      Caption         =   "�̰�����:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7800
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label �̰����� 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   9
      ToolTipText     =   "�D������"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label �D������ 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "�D������"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label �D����V 
      BackColor       =   &H00FFFF00&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "�D����V"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '�z��
      Caption         =   "�D����V:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label label2 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FF0000&
      BackStyle       =   0  '�z��
      Caption         =   "�D������:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label ������ñ 
      BackStyle       =   0  '�z��
      Caption         =   "�Х�2,4,6,8�䱱���V,5��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label �M�צW�� 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�g�Y�D"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label �Z�ũm�W 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "���:Form3"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2265
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***�`�N��޸��᭱�O�{�����ѻ����ΡA�b�u���Ҹռg�{���O�����g��
'�O�o���檺<KeyPreview>�]��<True>

Dim ���I�j�p As Integer   '�ŧi�o���ܼơA�Ϯ��I�C���Ҧ��j�p�β��ʶZ���ۦP�A���γo�ӭ�

Private Sub Form_KeyPress(KeyAscii As Integer)   '���H���U����A�n�ϳo�ƥ�o�͡A�O�o���檺KeyPreview�]��True
Select Case Chr(KeyAscii)                       '������U�ɷ|�ǤJ�ҫ����䪺���X�A�N�H�ǨӪ����X��P�_����
Case "2":                                        '�]���ǨӬO���X�A�gCHR()��ƴ�����J���r
   �D����V = "�U"                           '�p�G�ۦP�A�N���~���̫��F2�o����A�N���V�]��<�U>
Case "4":                                        '�P�z�P�_��L�Q�Ϊ��䤺�X�A���ۦP��
   �D����V = "��"                         '�N�]���۹諸�r�A�H�K���U�D���ʮɪ���V�̾ڡC
Case "6":
   �D����V = "�k"
Case "8":
   �D����V = "�W"
Case "5":
   �D����V = "��"
End Select
End Sub

Private Sub Form_Load()        '�{���@�}�l����ɡA�|�����檺�ƥ�A�ΨӰ��ܼƪ�ȳ]�w
Randomize                      '�]�����ϥΨ�<Rnd>�o��ơA�ҥH�n�[�o��A�Ϩ�üƧ����S���W�ߩ�
���I�j�p = �D(0).Width         '���ϩҦ��D�ΫC�쵥�j�A�ҥH����<�D(0).Width>�s�b�ŧi�L��<���I�j�p>�o�ܼƸ�
�D(0).Height = ���I�j�p        '��D�Y�����׳]����e�פ@�ˡA�o�˳D�Y�N�|�O�����
                              
�C��.Width = ���I�j�p          '��C��]�]�����j
�C��.Height = ���I�j�p

�D(0).Left = ��X��             '�Q�Τw�g�n�����I�ƨ�ơA�ӨϤ@�}�l<�D(0)>�H�N�b��椤��@�ӷs��m
�D(0).Top = ��Y��

�C��.Left = ��X��              '�P�˪��A��C��]�H����b��檺�Y�Ӯ��I�W
�C��.Top = ��Y��
End Sub

Private Function ��X��()                          '�ۼg���Ƶ{���A�ΨӦb��椤�H�����@�Ӯ��I�j�p�㭿�ƪ�X��m
�̤j���� = Int(Me.Width / ���I�j�p)             '�����e�װ����I�j�p�A��X�ثe��i��X���D
��X�� = Int(Rnd * �̤j����) * ���I�j�p             '�̩ү�񪺭��ƨ�Rnd�Ʀ���ƦA��<���I�j�p>�A�o�˫O�Ҩ��쪺��m�O�D�e������
End Function
Private Function ��Y��()                           '�ۼg���Ƶ{���A�ΨӦb��椤�H�����@�Ӯ��I�j�p�㭿�ƪ�Y��m
����ڰ��� = Me.Height - 600   '�`�N�OMe.Height���OWidth '�]����榳���Ŧ���D�ϡA�|���ά�600�����סA�һݴ,
�̤j���� = Int(����ڰ��� / ���I�j�p)             '������ڰ��סA�����I�j�p�A��X�ثe�a���i��X���D
��Y�� = Int(Rnd * �̤j����) * ���I�j�p             '���L�ѩ��檺Y�b�O���Ŧ���D�U���_�A�ҥH�n����@�ӭȡA�H�K��n��b�ݤ��쪺�a��
End Function                                       '�̩ү�񪺭��ƨ�Rnd�Ʀ���ƦA��<���I�j�p>�A�o�˫O�Ҩ��쪺��m�O�D�e������

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '����F���k�W��X�ɷ|�o�ͪ��ƥ�
Form4.Show                                  '�]���O�h�����A�ҥH�b����浲���e�n�Ұʿﶵ���C�Ҹծɦ��B�����g�C
End Sub                                     '�Ҹծɦ��q�����g

Private Sub �p�ɾ�_Timer()          '�T�w�ɶ��g�����|�Ӱ���@�M���ƥ�
If Me.WindowState = vbMinimized Then
   Exit Sub
End If
'�U���O����D�Y��m�O��A�H�K���D���{���ϥ�
�D�YX = �D(0).Left                      '���ʳD�Y�e����ثe��аO�bX,Y�ܼƸ̡A�H�K�᭱���D�����b�D�Y���
�D�YY = �D(0).Top

'�U���O�B�z<�D�Y����>���{��
Select Case �D����V                    '�̾�<�D����V>�ثe�e�{���r�ӨM�w�D�n�����Ӥ�V���@��<���I�j�p>�Z��
Case "�W":
   �D(0).Top = �D(0).Top - ���I�j�p     '�p�G�O�W�U�A�h���ܪ��O�D�Y���W��<�D(0).Top>�A�C�����ܤ@�ӳD���Z��
Case "�U":
   �D(0).Top = �D(0).Top + ���I�j�p
Case "��":
   �D(0).Left = �D(0).Left - ���I�j�p   '�p�G�O���k�A�h���ܪ��O�D�Y���W��<�D(0).Left>�A�C�����ܤ@�ӳD���Z��
Case "�k":
   �D(0).Left = �D(0).Left + ���I�j�p   '�������B�z�A�]���p�G�O���N�������ʡC
End Select

'�U���O�B�z<�D�Y��C��>���{��                             '�D�Y���ʧ��A�D���٨S�ʮɥ��P�_���S���Y��C��A�H�K����D���[��
If �D(0).Top = �C��.Top And �D(0).Left = �C��.Left Then   '�p�G�C��P�D(0)���W�ɸ򥪬ɬۦP�N��D�Y��C��F
   �C��.Left = ��X��                                      '�C��Q�r��A�N���C�촫�ӷs����m
   �C��.Top = ��Y��
   Load �D(�D������)                                      '�ƻs�@�ӳD��
   �D(�D������).Caption = "X"                              '�ϳD���W�S���r
   �D(�D������).BackColor = vbYellow                      '�ϳD����D�Y���P��
   �D(�D������).Visible = True                            '�Ϸs���ͪ��o�ӳD���X�{
   �p�ɾ�.Interval = �p�ɾ�.Interval * 0.8                '���F�W�[�ĪG�A�Y�@���C��N���D�[�t�ʤ���20
   �D������ = Val(�D������) + 1                           '�ç�<�D������>�[�@�A�H�K�U���Y��C��ɷ|�e�{�U�@�`
   If Val(�̰�����) < Val(�D������) Then                  '<�̰�����>�P<>���O�r��A�ҥH�@�w�n����val�Ʀ��ƭȨӤ��
      �̰����� = �D������                                 '�_�h�|�o��<"9">��<"10">�j���{�H�C
   End If                                                 '�i�O�n���ܬY���ܼƮɡA�������䪺�ܼƫo���i�[�Wval,�u�������ܼƦW��
End If


'�U���O�B�z<�D������>���{��
For i = 1 To �D������ - 1                                 '���ʳD���{���A���C�@�`����۫e�@�`��,�]���u�n�B�D���ҥH For �O�q 1 ��_���C
   �D��X = �D(i).Left                                     '�n���e�@��m�����e����ثe�y�Яd�U�bX2,Y2
   �D��Y = �D(i).Top
   �D(i).Left = �D�YX                                     '�u����ثe�o�D����D�Y��m
   �D(i).Top = �D�YY
   �D�YX = �D��X                                          '�w��X,Y��m�]���貾���o�`�D����m�A�ϤU�@�`�D���~�H��
   �D�YY = �D��Y                                          '�L�e�������`�N�O�D�Y�A�o�˴N�|�@�`��ۤ@�`��
Next

'�U���O�B�z<�D����>�{��
If �D(0).Left < 0 Or �D(0).Left > Me.Width Or �D(0).Top < 0 Or �D(0).Top > Me.Height - 600 Then '�P�_�O�_����
   MsgBox "����F"                       '�p�G�D(0)�����ɶ]���檺���k�~���γD(0)���W�ɶ]����W�U���~��
   �D(0).Left = ��X��                    '�N����F�A�o�ɦ�����ƭn���A�H��_�D���_�l���A�C
   �D(0).Top = ��Y��                     '�@.�q�ӰT�� ��msgbox���O�A
   �D����V = "��"               '�G.���D(0)��@�ӷs����
   For i = 1 To �D������ - 1             '�T.�]�����ǳD���w�e�{�A�A��Ҧ��D����������
      Unload �D(i)                       '�|.��D�����ק令�򥻪��סA�N�O1�`
   Next
   �D������ = 1                  '��.�]�����w��A�Ȫ����٨S�ݲM���A�N����F�A�ҥH<�D����V>�]��"��"�C
   �p�ɾ�.Interval = 300                 '��.��_��l���ʳt�סA�H�K�����A�@�}�l�N�ܧ�
End If
'�U���O�B�z<�r�D�ۺ�>�{��
If �D����V <> "��" Then                  '�P�_�O�_�D���r��ۤv�A���L�����������b<��>�����A���A�_�h��<��>�N�r��F
   ���r��ۤv = False
   For i = 1 To �D������ - 1             '��ڦ��D�����׳D�Y����A�ҥH�u�P�_��<�D������>��@�`
      If �D(0).Left = �D(i).Left And �D(0).Top = �D(i).Top Then  '�p�G<�D(0)>��D���Y�`<�D(J)>�����ɸ�W�ɬۦP�N�N��r��ۤv�F
         ���r��ۤv = True
      End If
   Next
   If ���r��ۤv = True Then
      MsgBox "�n�h"                   '�r��ۤv��e��������@�˭n������ƨӭ��s�Ұʳg�Y�D
      �D(0).Left = ��X��
      �D(0).Top = ��Y��
      �D����V = "��"
      For i = 1 To �D������ - 1
         Unload �D(i)
      Next
      �D������ = 1
      �p�ɾ�.Interval = 300           '�]���O�@�`�`�ˬd�D���O�_����D�Y�b�P��m�A���r��ۤv�ɷ|��D������
   End If
End If
End Sub
