VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�q���q��"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10440
   Icon            =   "�q���q��.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   10440
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton �s�� 
      BackColor       =   &H000080FF&
      Caption         =   "�s��"
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton ���� 
      BackColor       =   &H0000FFFF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox �ɦW 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6180
      TabIndex        =   1
      Text            =   "C:\TTT.TXT"
      Top             =   6600
      Width           =   3495
   End
   Begin VB.CommandButton �y�� 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   0
      Left            =   840
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label �Z�Ůy�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "���:Form4"
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
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "�s���ɦW:"
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
      Height          =   435
      Left            =   4140
      TabIndex        =   3
      Top             =   6600
      Width           =   1905
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   10320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label �M�׻��� 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�q���q��{��"
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
      Height          =   480
      Left            =   5088
      TabIndex        =   2
      Top             =   120
      Width           =   2736
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      X1              =   240
      X2              =   10260
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�q���q��{���A�D�n���� 1.���i�}����A2�q��A3�s�ɡA4���ɦ@�|�Ӱʧ@

Private Sub Form_Load()  '�{���}�l����ɡA�����i�}���󪺰ʧ@
For �� = 1 To 6          '�`�N<�y��>���s����Index�ݭn����0�A�Ϩ䦨���}�C
   For �C = 1 To 8       '�w�p���}�� 6 �ơA�C�� 8 �C
      Index = (�� - 1) * 8 + (�C - 1) '�Ĥ@�ƲĤ@�C<Index>��0,�Ĥ@�ƲĤG�C<Index>��1�A�ҥH�̭n�ت��Ƹ�C��X�۹諸Index��
      If Index > 0 Then   '�]����0�ӡA�O���ӴN���A�ҥH��0�٤j�����n�ƻs�X��
         Load �y��(Index)
      End If
      �y��(Index).Left = �y��(0).Left + �y��(0).Width * (�� - 1) '�ק�s�ͥX�Ӫ�<Left>�ȡA�Ĥ@�Ƹ�y��(0)�P��m�A�ĤG�ƭn�k���@�Ӯy��(0)���e�סA�P�z�ĤT�ƭn����Ӯy��(0)���e��
      �y��(Index).Top = �y��(0).Top - �y��(0).Height * (�C - 1)  '�ק�s�ͥX�Ӫ�<Top>�ȡA�Ĥ@�Ƹ�y��(0)�P��m�A�ĤG�ƭn�W���@�Ӯy��(0)�����סA�P�z�ĤT�ƭn����Ӯy��(0)�����סA�]���W�A�ҥH�n�δ�AY�b�V�W�V�p
      �y��(Index).Caption = (Index + 1)            '�b�s�ͪ�����W�[�W���D�A���D�N�O�s���[1�A�]���q���ߺD�q0��_�A�H�ߺD�̥����@�ӬO1
      �y��(Index).BackColor = vbYellow             '�����]������(�N���q��)�A��Q�I�ɦA�令���(�N��w�q��)
      �y��(Index).Visible = True                   '�s�ͪ����󤺩w�ݤ����A�ҥH�̫�n��<Visible>�אּ<True>�_�h�û��u�|�X�{�@�ӫ��s
   Next                  '�C��<For>���۹�y�N�O<Next>
Next            '�C��<For>���۹�y�N�O<Next>�A�����For,�ҥH�۹��m�]�n���Next
End Sub

Private Sub �y��_Click(Index As Integer)  '�q��O�� toggle �椬���檺�覡
If �y��(Index).BackColor = vbYellow Then   '�p�G���s���Q�I��O����(�N���q��)�A�I�F�n�令�q��
   �y��(Index).BackColor = vbGreen         '���I�����令���I��(�N��w�q��)�]���O�}�C����A�ҥH�I���@�ӳ��|�i���{���A���|���I�����Ӫ�Index�N�J�A�ҥH�����ק�Index���Ӫ��ݩʡA�N�O���I�������@�ӫ��s
   �y��(Index).Caption = (Index + 1) & ":�w�q��" '�ç�ӫ��s�e�{���r�令<�y��>�[�W":�w�q��"���r�A�`�N�s����r�����n�[ & �ӳs��
Else                                       '�p�G���O����A�N���I���ɭԬO���(�N��w�q��A�ҥH�n�令���q��)
   �y��(Index).BackColor = vbYellow        '��I����令��
   �y��(Index).Caption = (Index + 1)       '��W���e�{���r�����u��<�y��>
End If
End Sub

Private Sub �s��_Click()   '�s�ɶ��`�N�ϥΫe�n�}�ɡA�����ɭn�O�o���ɡA�_�h�������ĤG��
Open �ɦW.Text For Output As #1  '�}�ɮ榡�O Open(�}��) �ɦW For(���F) Output(��X) �̫�n���ӽs���A�@�볣�O#1�A�p�P�ɶ}�ɡA�ĤG�ӴN�s#2
For Index = 0 To 47              '��48�ӫ��s�{�p��X���ɮ�
   If �y��(Index).BackColor = vbGreen Then '�ˬd�ӫ��s���I����A�p�G�O<���>�N��w�q��A�_�h�N��<���q��>
      Print #1, "�w�q��"         '�`�ǧ�U���s�{�p��X����r�ɡA��X�����O�O Print #1
   Else
      Print #1, "���q��"         '�S�q��]�n��X�A�u����X���ɴN�|��48���ơA��n�̧ǹ�M��48���s�{�p
   End If
Next
Close #1                         '��X���A�n���W���ɡA�ê`�N<�ɦW>�o����n���@�Ӧ��Ī��ɦW�A�駹��n��Ͼ��ݬݿ�X���G�O�_���T
MsgBox "�s�ɧ���", vbInformation '�]��X�ܧ֡A�ϥΪ̱`���M�����S����X�A�ҥH��������ܤ@��<�s�ɧ���>���T���ص��ϥΪ̬�
End Sub

Private Sub ����_Click()
Open �ɦW.Text For Input As #1 '���ɸ�s�ɫܹ��A�u�O�ϦV�ʧ@�A�ҥH<For>�᭱�����O<Input>
For Index = 0 To 47            '�s48�Ӹ�ơA��M�N�n���^48�Ӹ��
   Line Input #1, AAA          'Ū�^��ƫ��O�� Line Input #1�A ����b�@���{���ܼƸ�
   If AAA = "�w�q��" Then      '���ɮפ�Ū�J�@�檺���e�A�Ӥ���A�]���@��O�@�Ӯy�쪺��ơA
      �y��(Index).BackColor = vbGreen                '�ҥH�p�G�r�O<�w�q��>�A�N���y���ӬO���q��A�`�N���B���r�A�n�M�s�ɮɰe�X���r�����ۦP
      �y��(Index).Caption = (Index + 1) & ":�w�q��"  '�⦹�y�쪺�I���⴫��<���>�A�ç���D�令�y���[<:�w�q��>
   Else
      �y��(Index).BackColor = vbYellow       '�p�G��藍�O<�w�q��>���N�N���q��A�ҥH��������I���Υu��ܮy��
      �y��(Index).Caption = (Index + 1)
   End If
Next
Close #1                     '���^�����A�ߧY���ɡA�ϥi�H����U�����s�ɩΨ���
MsgBox "���^����"            '��ܤ@�y�T���A���ϥΪ̪��D��Ƥw���^�A�`�N Msgbox�O���O�A�ҥHMsgbox�������O <����>�A�O<�ť�>�A�M��~�O�n��ܪ� "�r��"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '����F���k�W��X�ɷ|�o�ͪ��ƥ�
Form5.Show                                  '�]���O�h�����A�ҥH�b����浲���e�n�Ұʿﶵ���C�Ҹծɦ��B�����g�C
End Sub                                     '�Ҹծɦ��q�����g

