VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   Icon            =   "�s�����.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   8385
   ScaleWidth      =   9420
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton ���T 
      BackColor       =   &H0080C0FF&
      Caption         =   "�g�Y�D�C��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   6
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton ���@ 
      BackColor       =   &H0080C0FF&
      Caption         =   "���a���C��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton ���G 
      BackColor       =   &H0080C0FF&
      Caption         =   "���n�l�C��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   4
      Top             =   2700
      Width           =   2295
   End
   Begin VB.ListBox �D�� 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   2010
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   4440
      Width           =   6855
   End
   Begin VB.ListBox �D�� 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1230
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   1020
      Width           =   6855
   End
   Begin VB.ListBox �D�� 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1815
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   2460
      Width           =   6855
   End
   Begin VB.Label �X�p���� 
      AutoSize        =   -1  'True
      Caption         =   "���C���W�٥i�ժ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   12
      Top             =   7020
      Width           =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "�T�ӹC���U�ۿW�߼g�@�ӱM�ת��A�����L��"
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
      Left            =   300
      TabIndex        =   11
      Top             =   7500
      Width           =   9015
   End
   Begin VB.Label �p�p 
      Alignment       =   1  '�a�k���
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   10
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label �p�p 
      Alignment       =   1  '�a�k���
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label �p�p 
      Alignment       =   1  '�a�k���
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Index           =   2
      Left            =   1620
      TabIndex        =   8
      Top             =   5580
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "���C���W�٥i�ժ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "�q���ҵ{VB������"
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
      Left            =   2565
      TabIndex        =   0
      Top             =   480
      Width           =   3825
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub �M�׻���_Click()

End Sub

Private Sub Form_Load()
i = 0
�D��(i).AddItem "10��:�a���C0.5��|���@����m"
�D��(i).AddItem " 5��:�C���@�U�A<�w������>�Ʀr�|�[�@"
�D��(i).AddItem " 5��:�C�����a���@�U�A<��������>�Ʀr�|�[�@"
�D��(i).AddItem " 5��:<�����v>�|�̨C���@�U�O�����T��v�A�æ��p�ƨ�쪺��T��"
�D��(i).AddItem " 5��:�˼ƭp�ɾ��|�C��˼Ƥ@�A��s�|����p��"
�D��(i).AddItem "10��:�˼ư���p�ɫ�|�ݬO�_�A���@��,�åi�A���ε���"
i = i + 1
�D��(i).AddItem " 5��:�n�l�i�H���k�۰ʲ���"
�D��(i).AddItem " 5��:�Q�~�i�H���k�۰ʲ���"
�D��(i).AddItem " 5��:�a�����u�|��ƹ��b�a�����k��"
�D��(i).AddItem " 5��:���@�U�ƹ��A���u�|�_�����W"
�D��(i).AddItem " 5��:�C�o�g�@�ӭ��u<�o�g�ƶq>���Ʀr�|�[�@"
�D��(i).AddItem " 5��:���X���W��<���A>�|�X�{<�l�����u�@�T>�æ��^���u"
�D��(i).AddItem "10��:���u�����n�l<���A>�|�X�{<�����ؼ�>�æ��^���u"
�D��(i).AddItem " 5��:�C�����n�l<�R���ƶq>���Ʀr�|�[�@"
�D��(i).AddItem " 5��:���u�����Q�~<���A>�|�X�{<�����ؼ�>�æ��^���u"
i = i + 1
�D��(i).AddItem "10��:���Ʀr2,4,6,8,5<�D����V>�W���r�|�̧��ܦ��U,��,�k,�W,����"
�D��(i).AddItem " 5��:����ɳD�ΫC�쳣��۰��ܦ�<�D(0).Width>���ۦP�����"
�D��(i).AddItem " 5��:�}�l����ɨC���D�ΫC�쳣�i�H�b���W���P��m�X�{"
�D��(i).AddItem " 5��:�D�Y�|���<�D����V>����"
�D��(i).AddItem " 5��:�D�|���T���Y��C��A�C��|����m�A�D�|�[���@�`�A�t���ܧ�"
�D��(i).AddItem "10��:�D���|���T����۳D�Y����"
�D��(i).AddItem " 5��:�D�Y��C��A<�̰�����>�|���T��s�C"
�D��(i).AddItem " 5��:�D������@����|�X�{<����F>�B���T���^�D�A�i���s�A���@��"
�D��(i).AddItem "10��:�D�p�G�r��ۤv�|�X�{<�n�h>�B���T���^�D�A�i���s�A���@��"
���� = 0
For j = 0 To 2
   For i = 0 To �D��(j).ListCount - 1
      �p�p(j) = �p�p(j) + Val(�D��(j).List(i))
   Next
   ���� = ���� + �p�p(j)
Next
�X�p���� = "�X�p " & ���� & " ���A�ή欰60���A�W�L90�������C5����1���A�����Ҥή�����Һ�ή�C"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

End

End Sub

Private Sub ���@_Click()
Form1.Show
Unload Form2
Unload Form3
Form4.Hide

End Sub

Private Sub ���G_Click()
Unload Form1
Form2.Show
Unload Form3
Form4.Hide

End Sub

Private Sub ���T_Click()
Unload Form1
Unload Form2
Form3.Show
Form4.Hide

End Sub

