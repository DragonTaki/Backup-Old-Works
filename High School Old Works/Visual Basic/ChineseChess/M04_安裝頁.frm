VERSION 5.00
Begin VB.Form �w�˭� 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H�� - InstallShield Wizard"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15075
   Icon            =   "M04_�w�˭�.frx":0000
   LinkMode        =   1  '�ӷ�
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   15075
   StartUpPosition =   2  '�ù�����
   Visible         =   0   'False
   Begin VB.Frame �Ϧw�ˮ� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Timer �Ϧw�ˮحp�ɾ� 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7080
         Top             =   0
      End
      Begin VB.Frame �Ϧw�ˮث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   6120
         TabIndex        =   107
         Top             =   5760
         Width           =   1215
         Begin VB.Label �Ϧw�ˮب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   30
            TabIndex        =   108
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape �Ϧw�ˮب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox �Ϧw�ˮ�InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "@Gulim"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   106
         Text            =   "InstallShield"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label �Ϧw�ˮة��h�I�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -120
         TabIndex        =   105
         Top             =   5400
         Width           =   7815
      End
      Begin VB.Shape �Ϧw�ˮر��� 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '���z��
         BorderStyle     =   0  '�z��
         Height          =   270
         Index           =   0
         Left            =   300
         Shape           =   4  '�ꨤ�x��
         Top             =   2550
         Width           =   135
      End
      Begin VB.Shape �Ϧw�ˮر����� 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         Shape           =   4  '�ꨤ�x��
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label �Ϧw�ˮز��� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   104
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label �H�Ѧw�˵{�����b�����{�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�H�� �w�˵{�����b�����{���C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   103
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label �����H�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "���� �H�ѡC"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   102
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label ���� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   101
         Top             =   240
         Width           =   495
      End
      Begin VB.Label �Ϧw�ˮحI�� 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame �w�˪��A�� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Frame �w�˪��A�ث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   6120
         TabIndex        =   75
         Top             =   5760
         Width           =   1215
         Begin VB.Label �w�˪��A�ب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   30
            TabIndex        =   76
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape �w�˪��A�ب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Timer �w�˪��A�حp�ɾ� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7080
         Top             =   0
      End
      Begin VB.TextBox �w�˪��A��InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "@Gulim"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   67
         Text            =   "InstallShield"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Shape �w�˪��A�ر��� 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '���z��
         BorderStyle     =   0  '�z��
         Height          =   270
         Index           =   0
         Left            =   300
         Shape           =   4  '�ꨤ�x��
         Top             =   2550
         Width           =   135
      End
      Begin VB.Shape �w�˪��A�ر����� 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         Shape           =   4  '�ꨤ�x��
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label �w�˪��A 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�w�˪��A"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   71
         Top             =   600
         Width           =   855
      End
      Begin VB.Label �H�Ѧw�˵{�����b����ҭn�D���w�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�H�� �w�˵{�����b����ҭn�D���w�ˡC"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label �w�˪��A�ئw�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�w��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   69
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label �w�˪��A�ظ��| 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "C:\"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   2280
         Width           =   6735
      End
      Begin VB.Label �w�˪��A�حI�� 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label �w�˪��A�ة��h�I�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -240
         TabIndex        =   73
         Top             =   5400
         Width           =   7815
      End
   End
   Begin VB.Frame �ק�� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Caption         =   "s"
      Height          =   6300
      Left            =   -120
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.DriveListBox Drive 
         Height          =   300
         Left            =   240
         TabIndex        =   98
         Top             =   4920
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.PictureBox �ק�حק�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BorderStyle     =   0  '�S���ؽu
         FillColor       =   &H00FF0000&
         FillStyle       =   0  '���
         ForeColor       =   &H00FF0000&
         Height          =   2500
         Left            =   250
         ScaleHeight     =   2505
         ScaleWidth      =   3765
         TabIndex        =   54
         Top             =   1790
         Width           =   3760
         Begin VB.CheckBox �ק��MainApp 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "Main App"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   46
            Top             =   30
            Value           =   1  '�֨�
            Width           =   1095
         End
         Begin VB.Line �ק�ؽu 
            BorderStyle     =   3  '�I�u
            X1              =   120
            X2              =   360
            Y1              =   150
            Y2              =   150
         End
      End
      Begin VB.Frame �ק�ػ��� 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         ForeColor       =   &H80000002&
         Height          =   2655
         Left            =   4320
         TabIndex        =   53
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox �ק��InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "@Gulim"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Text            =   "InstallShield"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Frame �ק�ث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   3480
         TabIndex        =   42
         Top             =   5760
         Width           =   3855
         Begin VB.Label �ק�ب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   2670
            TabIndex        =   45
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label �ק�ؤU�@�B 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "�U�@�B(N)�@>"
            Height          =   255
            Left            =   1320
            TabIndex        =   44
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label �ק�ؤW�@�B 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "<�@�W�@�B(B)"
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape �ק�ؤW�@�B�� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape �ק�ؤU�@�B�� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape �ק�ب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   2640
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Label �ק�غϺвέp�Ѿl�Ŷ� 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�i�� 0.00 MB ���Ŷ�"
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   4680
         Width           =   3735
      End
      Begin VB.Label �ק�غϺвέp�ݭn�Ŷ� 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ݭn 29.1 MB ���Ŷ��]�b C �ϺФW�^"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   4440
         Width           =   3735
      End
      Begin VB.Shape �ק�ؤ���� 
         BorderColor     =   &H00FF8080&
         FillColor       =   &H00FFC0C0&
         Height          =   2550
         Left            =   240
         Top             =   1770
         Width           =   3795
      End
      Begin VB.Label �ק�ة��h�I�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -120
         TabIndex        =   52
         Top             =   5400
         Width           =   7815
      End
      Begin VB.Label ��ܤ��� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "��ܤ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
      Begin VB.Label ��ܦw�˵{���N�w�˪����� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "��ܦw�˵{���N�w�˪�����C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   49
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label �ק�ؤ�r 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "��ܱz�Q�n�w�˪�����A���襤�n����������C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label �ק�حI�� 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame �ק�״_������ 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Frame �ק�״_�����ث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   5760
         Width           =   3855
         Begin VB.Label �ק�״_�����ؤW�@�B 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "<�@�W�@�B(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label �ק�״_�����ؤU�@�B 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "�U�@�B(N)�@>"
            Height          =   255
            Left            =   1320
            TabIndex        =   35
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label �ק�״_�����ب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   2670
            TabIndex        =   34
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape �ק�״_�����ب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   2640
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape �ק�״_�����ؤW�@�B�� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   0
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape �ק�״_�����ؤU�@�B�� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox �ק�״_������InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "@Gulim"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Text            =   "InstallShield"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.OptionButton �ק�״_�β����{���ﶵ 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�״_(E)"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   3048
         Width           =   975
      End
      Begin VB.OptionButton �ק�״_�β����{���ﶵ 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����(R)"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   11
         Top             =   4176
         Width           =   975
      End
      Begin VB.OptionButton �ק�״_�β����{���ﶵ 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�ק�(M)"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   1920
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label ������r 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�����Ҧ��w�w�ˤ���C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label �״_��r 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "���s�w�˥H�e�w�˵{���Ҧw�˪��Ҧ��{������C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   3675
         Width           =   4095
      End
      Begin VB.Label �ק��r 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "��ܭn�s�W���s�{������ο�ܭn�������ثe�w�w�ˤ���C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   2550
         Width           =   5175
      End
      Begin VB.Label �ק�״_�����ؤ�r 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�w��ϥ� �H�� �w�˺��@�{���C�ϥΦ��{���i�H�ק�ثe���w�ˡC�Ы��@�U�U�C�@�ӿﶵ�C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   6135
      End
      Begin VB.Label �ק�״_�β����{�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�ק�B�״_�β����{���C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label �w�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�w��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.Label �ק�״_�����حI�� 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label �ק�״_�����ة��h�I�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -120
         TabIndex        =   16
         Top             =   5400
         Width           =   7815
      End
   End
   Begin VB.Frame �w�˧����� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   -120
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Frame �w�˧����ث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   3480
         TabIndex        =   92
         Top             =   5760
         Width           =   3855
         Begin VB.Label �w�˧����ب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2670
            TabIndex        =   95
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label �w�˧����ا��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   1395
            TabIndex        =   94
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape �w�˧����ا����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape �w�˧����ب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   2640
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label �w�˧����ؤW�@�B 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "<�@�W�@�B(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   0
            TabIndex        =   93
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape �w�˧����ؤW�@�B�� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   0
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox �w�˧�����r 
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   91
         Text            =   "M04_�w�˭�.frx":324A
         Top             =   1080
         Width           =   3255
      End
      Begin VB.PictureBox �w�˧�����InstallShield�� 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   0
         Picture         =   "M04_�w�˭�.frx":3273
         ScaleHeight     =   4155
         ScaleWidth      =   2490
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   0
         Width           =   2490
      End
      Begin VB.TextBox �w�˧�����InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "@Gulim"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   89
         Text            =   "InstallShield"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label �w�˧��� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�w�˧���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   97
         Top             =   240
         Width           =   855
      End
      Begin VB.Label �w�˧����حI�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -1200
         TabIndex        =   96
         Top             =   5400
         Width           =   8775
      End
   End
   Begin VB.Frame ���@������ 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   -120
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox ���@������InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "@Gulim"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   74
         Text            =   "InstallShield"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.PictureBox ���@������InstallShield�� 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   0
         Picture         =   "M04_�w�˭�.frx":80CF
         ScaleHeight     =   4155
         ScaleWidth      =   2490
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   0
         Width           =   2490
      End
      Begin VB.TextBox ���@������r 
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   61
         Text            =   "M04_�w�˭�.frx":CF2B
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Frame ���@�����ث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   3480
         TabIndex        =   57
         Top             =   5760
         Width           =   3855
         Begin VB.Label ���@�����ؤW�@�B 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "<�@�W�@�B(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label ���@�����ا��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   1395
            TabIndex        =   59
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label ���@�����ب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2670
            TabIndex        =   58
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape ���@�����ا����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape ���@�����ب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   2640
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape ���@�����ؤW�@�B�� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   0
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Label ���@�����حI�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -1200
         TabIndex        =   63
         Top             =   5400
         Width           =   8775
      End
      Begin VB.Label ���@���� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "���@����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   62
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame ������ 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   -120
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.PictureBox ������InstallShield�� 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   0
         Picture         =   "M04_�w�˭�.frx":CF5C
         ScaleHeight     =   4155
         ScaleWidth      =   2490
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   0
         Width           =   2490
      End
      Begin VB.Frame �����ث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   3480
         TabIndex        =   37
         Top             =   5760
         Width           =   3855
         Begin VB.Label �����ؤW�@�B 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "<�@�W�@�B(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label �����ب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2670
            TabIndex        =   40
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label �����ا��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   1395
            TabIndex        =   39
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape �����ا����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1380
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape �����ب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   2640
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape �����ؤW�@�B�� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   120
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox ������r 
         Appearance      =   0  '����
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   2640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "M04_�w�˭�.frx":11DB8
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label InstallShieldWizard���� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "InstallShield Wizard ����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label �����حI�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -1200
         TabIndex        =   20
         Top             =   5400
         Width           =   8775
      End
   End
   Begin VB.Frame ��s�� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox ��s��InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "@Gulim"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   80
         Text            =   "InstallShield"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Frame ��s�ث��s�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '�S���ؽu
         Height          =   375
         Left            =   6120
         TabIndex        =   78
         Top             =   5760
         Width           =   1215
         Begin VB.Label ��s�ب��� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "����"
            Height          =   255
            Left            =   30
            TabIndex        =   79
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape ��s�ب����� 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '���z��
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '�ꨤ�x��
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Shape ��s�ر����� 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         Shape           =   4  '�ꨤ�x��
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Shape ��s�ر��� 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '���z��
         BorderStyle     =   0  '�z��
         Height          =   270
         Index           =   0
         Left            =   300
         Shape           =   4  '�ꨤ�x��
         Top             =   2550
         Width           =   135
      End
      Begin VB.Label ��s 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "��s"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   87
         Top             =   240
         Width           =   495
      End
      Begin VB.Label ��s�ة��h�I�� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   -240
         TabIndex        =   86
         Top             =   5400
         Width           =   7815
      End
      Begin VB.Label ��s�ظ��| 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "C:\"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   2280
         Width           =   6735
      End
      Begin VB.Label ��s�ا�s 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "��s"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   83
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label �H�Ѧw�˵{�����b��s�̷ܳs���� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�H�� �w�˵{�����b��s�̷ܳs�����C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label ��s�H�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "��s �H�ѡC"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   81
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label ��s�حI�� 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame �ĤG�w�˭��� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   7560
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Timer ���b�w���ɮ׭p�ɾ� 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7080
         Top             =   0
      End
      Begin VB.CommandButton �ĤG�w�˭��ب��� 
         Caption         =   "����"
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton �ĤG�w�˭��ئw�� 
         Caption         =   "�w��"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label �ĤG�w�˭��ؽеy�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�еy�ԡA�w�˵{�����b�N �H�� �w�˨�z���q���W�C"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label �ĤG�w�˭��إ��b�w�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "���b�w��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Image �ĤG�w�˭��ع� 
         Height          =   3015
         Left            =   720
         Picture         =   "M04_�w�˭�.frx":11E53
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Shape ���� 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '���z��
         BorderStyle     =   0  '�z��
         Height          =   270
         Index           =   0
         Left            =   780
         Shape           =   4  '�ꨤ�x��
         Top             =   2070
         Width           =   135
      End
      Begin VB.Shape �ĤG�w�˭��ر����� 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   720
         Shape           =   4  '�ꨤ�x��
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label �ĤG�w�˭��إ��b�w���ɮצW�� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label �ĤG�w�˭��إ��b�w���ɮ� 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "���b�w���ɮ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label �ĤG�w�˭��حI�� 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '��u�T�w
         Height          =   1095
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame �Ĥ@�w�˭��� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   7560
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox ���� 
         Appearance      =   0  '����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "M04_�w�˭�.frx":1D0BE
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton �Ĥ@�w�˭��ؤU�@�B 
         Caption         =   "�w��"
         Height          =   375
         Left            =   3000
         TabIndex        =   0
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton �Ĥ@�w�˭��ب��� 
         Caption         =   "����"
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox �Ĥ@�w�˭��ئw�˻��� 
         Appearance      =   0  '����
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   109
         Text            =   "M04_�w�˭�.frx":1D0CA
         Top             =   480
         Width           =   6735
      End
   End
   Begin VB.Frame �ĤT�w�˭��� 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   6300
      Left            =   7560
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.CheckBox �}�ҶH�� 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2760
         TabIndex        =   32
         Top             =   5160
         Value           =   1  '�֨�
         Width           =   1695
      End
      Begin VB.CommandButton �ĤT�w�˭��ا��� 
         Caption         =   "����"
         Height          =   375
         Left            =   4920
         TabIndex        =   31
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Image �ĤT�w�˭��ع� 
         Height          =   4290
         Left            =   720
         Picture         =   "M04_�w�˭�.frx":1D196
         Stretch         =   -1  'True
         Top             =   360
         Width           =   5250
      End
   End
End
Attribute VB_Name = "�w�˭�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CountTimeInstallation As Long
Dim InstallCount As Long
Dim InstallNewPath As String
Dim InstallNewFilePath As String
Dim InstallOriginalPath As String
Private Sub Form_Load()
   �w�˭�.Height = 6700
   �w�˭�.Width = 7500
   �Ĥ@�w�˭���.Left = 0
   �Ĥ@�w�˭���.Top = 0
   �Ĥ@�w�˭���.Height = 12840
   �Ĥ@�w�˭���.Width = 7500
   �ĤG�w�˭���.Left = 0
   �ĤG�w�˭���.Top = 0
   �ĤG�w�˭���.Height = 12840
   �ĤG�w�˭���.Width = 7500
   �ĤT�w�˭���.Left = 0
   �ĤT�w�˭���.Top = 0
   �ĤT�w�˭���.Height = 12840
   �ĤT�w�˭���.Width = 7500
   �Ϧw�ˮ�.Left = 0
   �Ϧw�ˮ�.Top = 0
   �Ϧw�ˮ�.Height = 12840
   �Ϧw�ˮ�.Width = 7500
   ��s��.Left = 0
   ��s��.Top = 0
   ��s��.Height = 12840
   ��s��.Width = 7500
   ������.Left = 0
   ������.Top = 0
   ������.Height = 12840
   ������.Width = 7500
   �ק��.Left = 0
   �ק��.Top = 0
   �ק��.Height = 12840
   �ק��.Width = 7500
   ���@������.Left = 0
   ���@������.Top = 0
   ���@������.Height = 12840
   ���@������.Width = 7500
   �w�˧�����.Left = 0
   �w�˧�����.Top = 0
   �w�˧�����.Height = 12840
   �w�˧�����.Width = 7500
   �w�˪��A��.Left = 0
   �w�˪��A��.Top = 0
   �w�˪��A��.Height = 12840
   �w�˪��A��.Width = 7500
   �ק�״_������.Left = 0
   �ק�״_������.Top = 0
   �ק�״_������.Height = 12840
   �ק�״_������.Width = 7500
   '''''''''''''''''''''''''''''''''''''''''''''''''�P�_�O�_�w�ˤλݧ�s
   TheFile = "C:\WINDOWS\ChineseChess.ini"
   Results = Dir$(TheFile)
   If Results = "" Then '�ɮפ��s�b
      �Ĥ@�w�˭���.Visible = True
   Else '�ɮצs�b
      Open "C:\WINDOWS\ChineseChess.ini" For Input As #1 'Ū��
         Line Input #1, Temp
      Close #1
      If Val(Temp) < �}�l���.�ɮת����ƭ� Then '���ª����A�ݧ�s
         �w�˭�.Width = 7500
         ��s��.Visible = True
      ElseIf Val(Temp) = �}�l���.�ɮת����ƭ� Then '�����ۦP�A�ק�״_����
         �w�˭�.Width = 7500
         �ק�״_������.Visible = True
      Else '���s�����A����
         '
         �w�˭�.Width = 7500
      End If
   End If
   ����.Text = "�H�� " & �}�l���.�ɮת���
   �}�ҶH��.Caption = "�}�ҶH�� " & �}�l���.�ɮת���
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If ���@������.Visible = True Or ������.Visible = True Or �w�˧�����.Visible = True Then
      End
   ElseIf �Ϧw�ˮ�.Visible = True Then
         �Ϧw�ˮحp�ɾ�.Enabled = False
      Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
      If Ans = vbYes Then
         End
      Else
         Cancel = 1
         �Ϧw�ˮحp�ɾ�.Enabled = True
         Exit Sub
      End If
   End If
   ���b�w���ɮ׭p�ɾ�.Enabled = False
   Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
      If �ĤG�w�˭���.Visible = True Then
         ���b�w���ɮ׭p�ɾ�.Enabled = True
      End If
   End If
End Sub
Private Sub �Ϧw�ˮحp�ɾ�_Timer() ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CountTimeInstallation = CountTimeInstallation + 1
   If CountTimeInstallation = 1 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\00000404.016"
      �Ϧw�ˮر���(0).Visible = True
   ElseIf CountTimeInstallation = 10 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\00000404.256"
   ElseIf CountTimeInstallation = 20 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\age2_x1.exe"
   ElseIf CountTimeInstallation = 25 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\AGE2_X1.ICD"
      �Ϧw�ˮر���(1).Visible = True
      �Ϧw�ˮر���(2).Visible = True
   ElseIf CountTimeInstallation = 40 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\AOC10.exe"
      �Ϧw�ˮر���(3).Visible = True
   ElseIf CountTimeInstallation = 50 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\clcd16.dll"
      �Ϧw�ˮر���(4).Visible = True
   ElseIf CountTimeInstallation = 70 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\clcd32.dll"
   ElseIf CountTimeInstallation = 90 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\dplayerx.dll"
      �Ϧw�ˮر���(5).Visible = True
   ElseIf CountTimeInstallation = 95 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\drvmgt.dll"
   ElseIf CountTimeInstallation = 100 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EBUEula.dll"
      �Ϧw�ˮر���(6).Visible = True
   ElseIf CountTimeInstallation = 115 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\ebueulax.dll"
   ElseIf CountTimeInstallation = 120 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EBUSetup.sem"
      �Ϧw�ˮر���(7).Visible = True
   ElseIf CountTimeInstallation = 135 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\empires2.exe"
   ElseIf CountTimeInstallation = 200 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EULA.RTF"
      �Ϧw�ˮر���(8).Visible = True
   ElseIf CountTimeInstallation = 210 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EULAx.RTF"
      �Ϧw�ˮر���(9).Visible = True
   ElseIf CountTimeInstallation = 255 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\HA312W32.DLL"
   ElseIf CountTimeInstallation = 261 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\language.dll"
      �Ϧw�ˮر���(10).Visible = True
      �Ϧw�ˮر���(11).Visible = True
   ElseIf CountTimeInstallation = 270 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\language_x1.dll"
   ElseIf CountTimeInstallation = 296 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\language_x1_p1.dll"
      �Ϧw�ˮر���(12).Visible = True
   ElseIf CountTimeInstallation = 306 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\myth.acm"
   ElseIf CountTimeInstallation = 363 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player9.hki"
      �Ϧw�ˮر���(13).Visible = True
   ElseIf CountTimeInstallation = 379 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player10.hki"
   ElseIf CountTimeInstallation = 380 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player.nfo"
      �Ϧw�ˮر���(14).Visible = True
   ElseIf CountTimeInstallation = 382 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player.nfp"
   ElseIf CountTimeInstallation = 384 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player.nfx"
      �Ϧw�ˮر���(15).Visible = True
   ElseIf CountTimeInstallation = 386 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Readme.rtf"
   ElseIf CountTimeInstallation = 389 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Readmex.rtf"
      �Ϧw�ˮر���(16).Visible = True
   ElseIf CountTimeInstallation = 399 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\ReadmeX_a.rtf"
   ElseIf CountTimeInstallation = 410 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\scenariobkg.bmp"
      �Ϧw�ˮر���(17).Visible = True
   ElseIf CountTimeInstallation = 420 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SETUPENU.DLL"
   ElseIf CountTimeInstallation = 445 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SetupReg.exe"
   ElseIf CountTimeInstallation = 450 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SHW32.DLL"
      �Ϧw�ˮر���(18).Visible = True
   ElseIf CountTimeInstallation = 475 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\STPENUX.DLL"
   ElseIf CountTimeInstallation = 480 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\yi8oytru.dll"
      �Ϧw�ˮر���(19).Visible = True
   ElseIf CountTimeInstallation = 490 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\rirsegsda.dll"
   ElseIf CountTimeInstallation = 512 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\q35tyrwe.dll"
      �Ϧw�ˮر���(20).Visible = True
   ElseIf CountTimeInstallation = 515 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\aeghjda.dll"
   ElseIf CountTimeInstallation = 516 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\serhyjykf.dll"
      �Ϧw�ˮر���(21).Visible = True
   ElseIf CountTimeInstallation = 520 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\sahjsdgh.dll"
   ElseIf CountTimeInstallation = 570 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SaveGame\Save.dll"
      �Ϧw�ˮر���(22).Visible = True
   ElseIf CountTimeInstallation = 580 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SaveGame\Inf.inf"
      �Ϧw�ˮر���(23).Visible = True
   ElseIf CountTimeInstallation = 595 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SaveGame\Top.txt"
      �Ϧw�ˮر���(24).Visible = True
   ElseIf CountTimeInstallation = 600 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Sound\0003.mp3"
      �Ϧw�ˮر���(25).Visible = True
   ElseIf CountTimeInstallation = 610 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Sound\0006.mp3"
      �Ϧw�ˮر���(26).Visible = True
   ElseIf CountTimeInstallation = 620 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Sound\Credits.mp3"
      �Ϧw�ˮر���(27).Visible = True
   ElseIf CountTimeInstallation = 630 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\...\Sound\Revolootin.mp3"
      �Ϧw�ˮر���(28).Visible = True
   ElseIf CountTimeInstallation = 640 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\...\Sound\SomeOfAKind.mp3"
      �Ϧw�ˮر���(29).Visible = True
   ElseIf CountTimeInstallation = 660 Then '����
      Call Kill_File("C:\WINDOWS\ChineseChess.ini")
      Call Kill_File("C:\Program Files\ChineseChess\Chinese Chess 2.3.0.exe")
      �Ϧw�ˮحp�ɾ�.Enabled = False
      Temp = MsgBox("The ChineseChess Has Been Removed.", vbOKOnly + vbInformation, "")
      End
   End If
End Sub
Private Sub ���b�w���ɮ׭p�ɾ�_Timer()
   CountTimeInstallation = CountTimeInstallation + 1
   If CountTimeInstallation = 1 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\00000404.016"
      ����(0).Visible = True
   ElseIf CountTimeInstallation = 10 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\00000404.256"
   ElseIf CountTimeInstallation = 20 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\age2_x1.exe"
   ElseIf CountTimeInstallation = 25 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\AGE2_X1.ICD"
      ����(1).Visible = True
      ����(2).Visible = True
   ElseIf CountTimeInstallation = 40 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\AOC10.exe"
      ����(3).Visible = True
   ElseIf CountTimeInstallation = 50 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\clcd16.dll"
      ����(4).Visible = True
   ElseIf CountTimeInstallation = 70 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\clcd32.dll"
   ElseIf CountTimeInstallation = 90 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\dplayerx.dll"
      ����(5).Visible = True
   ElseIf CountTimeInstallation = 95 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\drvmgt.dll"
   ElseIf CountTimeInstallation = 100 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\EBUEula.dll"
      ����(6).Visible = True
   ElseIf CountTimeInstallation = 115 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\ebueulax.dll"
   ElseIf CountTimeInstallation = 120 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\EBUSetup.sem"
      ����(7).Visible = True
   ElseIf CountTimeInstallation = 135 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\empires2.exe"
   ElseIf CountTimeInstallation = 200 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\EULA.RTF"
      ����(8).Visible = True
   ElseIf CountTimeInstallation = 210 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\EULAx.RTF"
      ����(9).Visible = True
   ElseIf CountTimeInstallation = 255 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\HA312W32.DLL"
   ElseIf CountTimeInstallation = 261 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\language.dll"
      ����(10).Visible = True
      ����(11).Visible = True
   ElseIf CountTimeInstallation = 270 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\language_x1.dll"
   ElseIf CountTimeInstallation = 296 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\language_x1_p1.dll"
      ����(12).Visible = True
   ElseIf CountTimeInstallation = 306 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\myth.acm"
   ElseIf CountTimeInstallation = 363 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\player9.hki"
      ����(13).Visible = True
   ElseIf CountTimeInstallation = 379 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\player10.hki"
   ElseIf CountTimeInstallation = 380 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\player.nfo"
      ����(14).Visible = True
   ElseIf CountTimeInstallation = 382 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\player.nfp"
   ElseIf CountTimeInstallation = 384 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\player.nfx"
      ����(15).Visible = True
   ElseIf CountTimeInstallation = 386 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\Readme.rtf"
   ElseIf CountTimeInstallation = 389 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\Readmex.rtf"
      ����(16).Visible = True
   ElseIf CountTimeInstallation = 399 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\ReadmeX_a.rtf"
   ElseIf CountTimeInstallation = 410 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\scenariobkg.bmp"
      ����(17).Visible = True
   ElseIf CountTimeInstallation = 420 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\SETUPENU.DLL"
   ElseIf CountTimeInstallation = 445 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\SetupReg.exe"
   ElseIf CountTimeInstallation = 450 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\SHW32.DLL"
      ����(18).Visible = True
   ElseIf CountTimeInstallation = 475 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\STPENUX.DLL"
   ElseIf CountTimeInstallation = 480 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\yi8oytru.dll"
      ����(19).Visible = True
   ElseIf CountTimeInstallation = 490 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\rirsegsda.dll"
   ElseIf CountTimeInstallation = 512 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\q35tyrwe.dll"
      ����(20).Visible = True
   ElseIf CountTimeInstallation = 515 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\aeghjda.dll"
   ElseIf CountTimeInstallation = 516 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\serhyjykf.dll"
      ����(21).Visible = True
   ElseIf CountTimeInstallation = 520 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\sahjsdgh.dll"
   ElseIf CountTimeInstallation = 570 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\SaveGame\Save.dll"
      ����(22).Visible = True
   ElseIf CountTimeInstallation = 580 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\SaveGame\Inf.inf"
      ����(23).Visible = True
   ElseIf CountTimeInstallation = 595 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\SaveGame\Top.txt"
      ����(24).Visible = True
   ElseIf CountTimeInstallation = 600 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\Sound\0003.mp3"
      ����(25).Visible = True
   ElseIf CountTimeInstallation = 610 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\Sound\0006.mp3"
      ����(26).Visible = True
   ElseIf CountTimeInstallation = 620 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\Chinese Chess\Sound\Credits.mp3"
      ����(27).Visible = True
   ElseIf CountTimeInstallation = 630 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\...\Sound\Revolootin.mp3"
      ����(28).Visible = True
   ElseIf CountTimeInstallation = 640 Then
      �ĤG�w�˭��إ��b�w���ɮצW��.Caption = "C:\Program Files\...\Sound\SomeOfAKind.mp3"
      ����(29).Visible = True
   ElseIf CountTimeInstallation = 660 Then '�����w��
   ''
      Call RealInstall '�u���w��
   ''
      Open "C:\WINDOWS\ChineseChess.ini" For Output As #1 '�إ��ɮ�
         Print #1, �}�l���.�ɮת����ƭ� '�g�J������
      Close #1
      For i = 0 To 29
         ����(i).Visible = False
      Next
      ���b�w���ɮ׭p�ɾ�.Enabled = False
      �ĤG�w�˭���.Visible = False
      �ĤT�w�˭���.Visible = True
   End If
End Sub
Private Sub RealInstall() '�u���w�˵{����
   '�g�J�T�{��
   'Open "C:\WINDOWS\ChineseChess.ini" For Output As #1 '�إ��ɮ�
   '   Print #1, "V2.3" '�g�J������
   'Close #1
   '����~���{��
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\FilePath.dll")
   Open "C:\Program Files\MakeFilesPath.txt" For Output As #1 '�إ��ɮ�
      Print #1, "ChineseChess" '�ɮצW��
   Close #1
   Shell ("C:\Program Files\FilePath.dll"), vbHide
Return0:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '�{���B�椤
       GoTo Return0
   Else '�{�����B��
      Call Kill_File("C:\Program Files\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\FilePath.dll")
   End If
   '
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\ChineseChess\FilePath.dll")
   Open "C:\Program Files\ChineseChess\MakeFilesPath.txt" For Output As #1 '�إ��ɮ�
      Print #1, "bin" '�ɮצW��
   Close #1
   Shell ("C:\Program Files\ChineseChess\FilePath.dll"), vbHide
Return1:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '�{���B�椤
       GoTo Return1
   Else '�{�����B��
      Call Kill_File("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\ChineseChess\FilePath.dll")
   End If
   '
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\ChineseChess\FilePath.dll")
   Open "C:\Program Files\ChineseChess\MakeFilesPath.txt" For Output As #1 '�إ��ɮ�
      Print #1, "Picture" '�ɮצW��
   Close #1
   Shell ("C:\Program Files\ChineseChess\FilePath.dll"), vbHide
Return2:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '�{���B�椤
       GoTo Return2
   Else '�{�����B��
      Call Kill_File("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\ChineseChess\FilePath.dll")
   End If
   '
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\ChineseChess\FilePath.dll")
   Open "C:\Program Files\ChineseChess\MakeFilesPath.txt" For Output As #1 '�إ��ɮ�
      Print #1, "Sound" '�ɮצW��
   Close #1
   Shell ("C:\Program Files\ChineseChess\FilePath.dll"), vbHide
Return3:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '�{���B�椤
       GoTo Return2
   Else '�{�����B��
      Call Kill_File("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\ChineseChess\FilePath.dll")
   End If
   '�ƻs�ɮ�
   InstallCount = 0
DoNextInstall:
   InstallCount = InstallCount + 1
   'CanDoNext = False
   Call RealInstallCount
   �ĤG�w�˭��إ��b�w���ɮצW��.Caption = InstallNewPath & InstallNewFilePath
   FileCopy (App.Path & "\" & InstallOriginalPath), (InstallNewPath & "\" & InstallNewFilePath)
   '
   'For i = 1 To 100
   '   For j = 1 To 100
   '   CountTime = CountTime + 1
   '   Next
   'Next
   '
   'InstallTimer.Enabled = True
   'Call InstallTimer_Timer
   If InstallCount = 112 Then '�w�˧���
      GoTo EndInstall
   End If
'CheckTimer:
   'If CanDoNext = True Then
      GoTo DoNextInstall
   'Else
   '   GoTo CheckTimer
   'End If
EndInstall:
   '�ĤG�w�˭���.Visible = False
   '�ĤT�w�˭���.Visible = True
End Sub
Private Sub Kill_File(ByRef Dir_String As String)
On Error GoTo Exit_out
Kill Dir_String
Exit_out:
End Sub
Private Sub RealInstallCount() '�u���w�ˤ��e
   If InstallCount = 1 Then 'exe
      InstallOriginalPath = "\Chess.dll"
      InstallNewPath = "C:\Program Files\ChineseChess\" 'Exe Path
      InstallNewFilePath = "Chinese Chess 2.3.0.exe"
   ElseIf InstallCount = 2 Then 'mp3
      InstallOriginalPath = "\Sound\0357130187950903543.s"
      InstallNewPath = "C:\Program Files\ChineseChess\Sound\" 'Sound Path
      InstallNewFilePath = "0003.mp3"
   ElseIf InstallCount = 3 Then
      InstallOriginalPath = "\Sound\1359385674917984206.s"
      'InstallNewPath =
      InstallNewFilePath = "0006.mp3"
   ElseIf InstallCount = 4 Then
      InstallOriginalPath = "\Sound\9284639256293457289.s"
      'InstallNewPath =
      InstallNewFilePath = "Credits.mp3"
   ElseIf InstallCount = 5 Then
      InstallOriginalPath = "\Sound\5717890367893847564.s"
      'InstallNewPath =
      InstallNewFilePath = "Revolootin.mp3"
   ElseIf InstallCount = 6 Then
      InstallOriginalPath = "\Sound\2078403903470298529.s"
      'InstallNewPath =
      InstallNewFilePath = "SomeOfAKind.mp3"
   ElseIf InstallCount = 7 Then 'tmp
      InstallOriginalPath = "\Temp\xnviewbe.tmp"
      InstallNewPath = "C:\Program Files\ChineseChess\bin\" 'Tmp Path
      InstallNewFilePath = "xnviewbe.tmp"
   ElseIf InstallCount = 8 Then
      InstallOriginalPath = "\Temp\xnviewbg.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewbg.tmp"
   ElseIf InstallCount = 9 Then
      InstallOriginalPath = "\Temp\xnviewbr.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewbr.tmp"
   ElseIf InstallCount = 10 Then
      InstallOriginalPath = "\Temp\xnviewca.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewca.tmp"
   ElseIf InstallCount = 11 Then
      InstallOriginalPath = "\Temp\xnviewcs.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewcs.tmp"
   ElseIf InstallCount = 12 Then
      InstallOriginalPath = "\Temp\xnviewda.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewda.tmp"
   ElseIf InstallCount = 13 Then
      InstallOriginalPath = "\Temp\xnviewde.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewde.tmp"
   ElseIf InstallCount = 14 Then
      InstallOriginalPath = "\Temp\xnviewel.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewel.tmp"
   ElseIf InstallCount = 15 Then
      InstallOriginalPath = "\Temp\xnviewes.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewes.tmp"
   ElseIf InstallCount = 16 Then
      InstallOriginalPath = "\Temp\xnviewfr.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewfr.tmp"
   ElseIf InstallCount = 17 Then
      InstallOriginalPath = "\Temp\xnviewgl.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewgl.tmp"
   ElseIf InstallCount = 18 Then
      InstallOriginalPath = "\Temp\xnviewhe.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewhe.tmp"
   ElseIf InstallCount = 19 Then
      InstallOriginalPath = "\Temp\xnviewhr.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewhr.tmp"
   ElseIf InstallCount = 20 Then
      InstallOriginalPath = "\Temp\xnviewhu.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewhu.tmp"
   ElseIf InstallCount = 21 Then
      InstallOriginalPath = "\Temp\xnviewit.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewit.tmp"
   ElseIf InstallCount = 22 Then
      InstallOriginalPath = "\Temp\xnviewja.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewja.tmp"
   ElseIf InstallCount = 23 Then
      InstallOriginalPath = "\Temp\xnviewko.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewko.tmp"
   ElseIf InstallCount = 24 Then
      InstallOriginalPath = "\Temp\xnviewlt.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewlt.tmp"
   ElseIf InstallCount = 25 Then
      InstallOriginalPath = "\Temp\xnviewlv.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewlv.tmp"
   ElseIf InstallCount = 26 Then
      InstallOriginalPath = "\Temp\xnviewnl.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewnl.tmp"
   ElseIf InstallCount = 27 Then
      InstallOriginalPath = "\Temp\xnviewno.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewno.tmp"
   ElseIf InstallCount = 28 Then
      InstallOriginalPath = "\Temp\xnviewpl.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewpl.tmp"
   ElseIf InstallCount = 29 Then
      InstallOriginalPath = "\Temp\xnviewpt.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewpt.tmp"
   ElseIf InstallCount = 30 Then
      InstallOriginalPath = "\Temp\xnviewro.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewro.tmp"
   ElseIf InstallCount = 31 Then
      InstallOriginalPath = "\Temp\xnviewru.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewru.tmp"
   ElseIf InstallCount = 32 Then
      InstallOriginalPath = "\Temp\xnviewsk.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewsk.tmp"
   ElseIf InstallCount = 33 Then
      InstallOriginalPath = "\Temp\xnviewsl.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewsl.tmp"
   ElseIf InstallCount = 34 Then
      InstallOriginalPath = "\Temp\xnviewsr.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewsr.tmp"
   ElseIf InstallCount = 35 Then
      InstallOriginalPath = "\Temp\xnviewsv.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewsv.tmp"
   ElseIf InstallCount = 36 Then
      InstallOriginalPath = "\Temp\xnviewtr.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewtr.tmp"
   ElseIf InstallCount = 37 Then
      InstallOriginalPath = "\Temp\xnviewtw.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewtw.tmp"
   ElseIf InstallCount = 38 Then
      InstallOriginalPath = "\Temp\xnviewuk.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewuk.tmp"
   ElseIf InstallCount = 39 Then
      InstallOriginalPath = "\Temp\xnviewzh.tmp"
      'InstallNewPath =
      InstallNewFilePath = "xnviewzh.tmp"
   ElseIf InstallCount = 40 Then 'chi
      InstallOriginalPath = "\pic\0.che"
      InstallNewPath = "C:\Program Files\ChineseChess\Picture\" 'Chi Path
      InstallNewFilePath = "0.che"
   ElseIf InstallCount = 41 Then
      InstallOriginalPath = "\pic\1.che"
      'InstallNewPath =
      InstallNewFilePath = "1.che"
   ElseIf InstallCount = 42 Then
      InstallOriginalPath = "\pic\2.che"
      'InstallNewPath =
      InstallNewFilePath = "2.che"
   ElseIf InstallCount = 43 Then
      InstallOriginalPath = "\pic\3.che"
      'InstallNewPath =
      InstallNewFilePath = "3.che"
   ElseIf InstallCount = 44 Then
      InstallOriginalPath = "\pic\4.che"
      'InstallNewPath =
      InstallNewFilePath = "4.che"
   ElseIf InstallCount = 45 Then
      InstallOriginalPath = "\pic\5.che"
      'InstallNewPath =
      InstallNewFilePath = "5.che"
   ElseIf InstallCount = 46 Then
      InstallOriginalPath = "\pic\6.che"
      'InstallNewPath =
      InstallNewFilePath = "6.che"
   ElseIf InstallCount = 47 Then
      InstallOriginalPath = "\pic\7.che"
      'InstallNewPath =
      InstallNewFilePath = "7.che"
   ElseIf InstallCount = 48 Then
      InstallOriginalPath = "\pic\8.che"
      'InstallNewPath =
      InstallNewFilePath = "8.che"
   ElseIf InstallCount = 49 Then
      InstallOriginalPath = "\pic\9.che"
      'InstallNewPath =
      InstallNewFilePath = "9.che"
   ElseIf InstallCount = 50 Then
      InstallOriginalPath = "\pic\10.che"
      'InstallNewPath =
      InstallNewFilePath = "10.che"
   ElseIf InstallCount = 51 Then
      InstallOriginalPath = "\pic\11.che"
      'InstallNewPath =
      InstallNewFilePath = "11.che"
   ElseIf InstallCount = 52 Then
      InstallOriginalPath = "\pic\12.che"
      'InstallNewPath =
      InstallNewFilePath = "12.che"
   ElseIf InstallCount = 53 Then
      InstallOriginalPath = "\pic\13.che"
      'InstallNewPath =
      InstallNewFilePath = "13.che"
   ElseIf InstallCount = 54 Then
      InstallOriginalPath = "\pic\14.che"
      'InstallNewPath =
      InstallNewFilePath = "14.che"
   ElseIf InstallCount = 55 Then
      InstallOriginalPath = "\pic\15.che"
      'InstallNewPath =
      InstallNewFilePath = "15.che"
   ElseIf InstallCount = 56 Then
      InstallOriginalPath = "\pic\16.che"
      'InstallNewPath =
      InstallNewFilePath = "16.che"
   ElseIf InstallCount = 57 Then
      InstallOriginalPath = "\pic\17.che"
      'InstallNewPath =
      InstallNewFilePath = "17.che"
   ElseIf InstallCount = 58 Then
      InstallOriginalPath = "\pic\18.che"
      'InstallNewPath =
      InstallNewFilePath = "18.che"
   ElseIf InstallCount = 59 Then
      InstallOriginalPath = "\pic\19.che"
      'InstallNewPath =
      InstallNewFilePath = "19.che"
   ElseIf InstallCount = 60 Then
      InstallOriginalPath = "\pic\20.che"
      'InstallNewPath =
      InstallNewFilePath = "20.che"
   ElseIf InstallCount = 61 Then
      InstallOriginalPath = "\pic\21.che"
      'InstallNewPath =
      InstallNewFilePath = "21.che"
   ElseIf InstallCount = 62 Then
      InstallOriginalPath = "\pic\22.che"
      'InstallNewPath =
      InstallNewFilePath = "22.che"
   ElseIf InstallCount = 63 Then
      InstallOriginalPath = "\pic\23.che"
      'InstallNewPath =
      InstallNewFilePath = "23.che"
   ElseIf InstallCount = 64 Then
      InstallOriginalPath = "\pic\24.che"
      'InstallNewPath =
      InstallNewFilePath = "24.che"
   ElseIf InstallCount = 65 Then
      InstallOriginalPath = "\pic\25.che"
      'InstallNewPath =
      InstallNewFilePath = "25.che"
   ElseIf InstallCount = 66 Then
      InstallOriginalPath = "\pic\26.che"
      'InstallNewPath =
      InstallNewFilePath = "26.che"
   ElseIf InstallCount = 67 Then
      InstallOriginalPath = "\pic\27.che"
      'InstallNewPath =
      InstallNewFilePath = "27.che"
   ElseIf InstallCount = 68 Then
      InstallOriginalPath = "\pic\28.che"
      'InstallNewPath =
      InstallNewFilePath = "28.che"
   ElseIf InstallCount = 69 Then
      InstallOriginalPath = "\pic\29.che"
      'InstallNewPath =
      InstallNewFilePath = "29.che"
   ElseIf InstallCount = 70 Then
      InstallOriginalPath = "\pic\30.che"
      'InstallNewPath =
      InstallNewFilePath = "30.che"
   ElseIf InstallCount = 71 Then
      InstallOriginalPath = "\pic\31.che"
      'InstallNewPath =
      InstallNewFilePath = "31.che"
   ElseIf InstallCount = 72 Then
      InstallOriginalPath = "\pic\32.che"
      'InstallNewPath =
      InstallNewFilePath = "32.che"
   ElseIf InstallCount = 73 Then 'fil
      InstallOriginalPath = "\ChineseChess.fil"
      InstallNewPath = "C:\WINDOWS\" 'ChineseChess.ini Path
      InstallNewFilePath = "ChineseChess.ini"
   ElseIf InstallCount = 74 Then 'dat
      InstallOriginalPath = "\Temp\Cpuinf32.dat"
      InstallNewPath = "C:\Program Files\ChineseChess\" 'Dat Path
      InstallNewFilePath = "Cpuinf32.dat"
   ElseIf InstallCount = 75 Then 'dll
      InstallOriginalPath = "\Temp\HerWizard.dll"
      InstallNewPath = "C:\Program Files\ChineseChess\bin" 'Bin Path
      InstallNewFilePath = "HerWizard.bin"
   ElseIf InstallCount = 76 Then
      InstallOriginalPath = "\Temp\HerWizardRC.dll"
      'InstallNewPath =
      InstallNewFilePath = "HerWizardRC.bin"
   ElseIf InstallCount = 77 Then
      InstallOriginalPath = "\Temp\HerWizardTask.dll"
      'InstallNewPath =
      InstallNewFilePath = "HerWizardTask.bin"
   ElseIf InstallCount = 78 Then
      InstallOriginalPath = "\Temp\HerWizPGEdit.dll"
      'InstallNewPath =
      InstallNewFilePath = "HerWizPGEdit.bin"
   ElseIf InstallCount = 79 Then
      InstallOriginalPath = "\Temp\HerWizPGImport.dll"
      'InstallNewPath =
      InstallNewFilePath = "HerWizPGImport.bin"
   ElseIf InstallCount = 80 Then
      InstallOriginalPath = "\Temp\HerWizPGOutput.dll"
      'InstallNewPath =
      InstallNewFilePath = "HerWizPGOutput.bin"
   ElseIf InstallCount = 81 Then
      InstallOriginalPath = "\Temp\HerWizProject.dll"
      'InstallNewPath =
      InstallNewFilePath = "HerWizProject.bin"
   ElseIf InstallCount = 82 Then 'Fake mp3
      InstallOriginalPath = "\Sound\5646156761316764379.s"
      InstallNewPath = "C:\Program Files\ChineseChess\Sound\" 'Sound Path
      InstallNewFilePath = "0001.mp3"
   ElseIf InstallCount = 83 Then
      InstallOriginalPath = "\Sound\1345136210342913136.s"
      'InstallNewPath =
      InstallNewFilePath = "0002.mp3"
   ElseIf InstallCount = 84 Then
      InstallOriginalPath = "\Sound\0497894655942568964.s"
      'InstallNewPath =
      InstallNewFilePath = "0004.mp3"
   ElseIf InstallCount = 85 Then
      InstallOriginalPath = "\Sound\3546806463765419799.s"
      'InstallNewPath =
      InstallNewFilePath = "0005.mp3"
   ElseIf InstallCount = 86 Then
      InstallOriginalPath = "\Sound\5734444938443613231.s"
      'InstallNewPath =
      InstallNewFilePath = "0007.mp3"
   ElseIf InstallCount = 87 Then
      InstallOriginalPath = "\Sound\1684985614651623179.s"
      'InstallNewPath =
      InstallNewFilePath = "0008.mp3"
   ElseIf InstallCount = 88 Then 'Fake dll
      InstallOriginalPath = "\Sound\uiplA6.dll"
      'InstallNewPath =
      InstallNewFilePath = "uiplA6.dat"
   ElseIf InstallCount = 89 Then
      InstallOriginalPath = "\Sound\uiplM5.dll"
      'InstallNewPath =
      InstallNewFilePath = "uiplM5.dat"
   ElseIf InstallCount = 90 Then
      InstallOriginalPath = "\Sound\uiplM6.dll"
      'InstallNewPath =
      InstallNewFilePath = "uiplM6.dat"
   ElseIf InstallCount = 91 Then
      InstallOriginalPath = "\Sound\uiplP6.dll"
      'InstallNewPath =
      InstallNewFilePath = "uiplP6.dat"
   ElseIf InstallCount = 92 Then
      InstallOriginalPath = "\Sound\uiplPX.dll"
      'InstallNewPath =
      InstallNewFilePath = "uiplPX.dat"
   ElseIf InstallCount = 93 Then
      InstallOriginalPath = "\Sound\uiplW7.dll"
      'InstallNewPath =
      InstallNewFilePath = "uiplW7.dat"
   ElseIf InstallCount = 94 Then 'Fake dll
      InstallOriginalPath = "\Temp\PS_ArchiveBuilder.dll"
      InstallNewPath = "C:\Program Files\ChineseChess\"
      InstallNewFilePath = "ArchiveBuilder.dll"
   ElseIf InstallCount = 95 Then
      InstallOriginalPath = "\Temp\PS_ArchiveRC.dll"
      'InstallNewPath =
      InstallNewFilePath = "ArchiveRC.dll"
   ElseIf InstallCount = 96 Then
      InstallOriginalPath = "\Temp\PS_CacheBufMgr.dll"
      'InstallNewPath =
      InstallNewFilePath = "CacheBufMgr.dll"
   ElseIf InstallCount = 97 Then
      InstallOriginalPath = "\Temp\PS_CommonControl.dll"
      'InstallNewPath =
      InstallNewFilePath = "CommonControl.dll"
   ElseIf InstallCount = 98 Then
      InstallOriginalPath = "\Temp\PS_CommonControlRC.dll"
      'InstallNewPath =
      InstallNewFilePath = "CommonControlRC.dll"
   ElseIf InstallCount = 99 Then
      InstallOriginalPath = "\Temp\PS_GUIBase.dll"
      'InstallNewPath =
      InstallNewFilePath = "GUIBase.dll"
   ElseIf InstallCount = 100 Then
      InstallOriginalPath = "\Temp\PS_ITxFrMapMgr.dll"
      'InstallNewPath =
      InstallNewFilePath = "ITxFrMapMgr.dll"
   ElseIf InstallCount = 101 Then
      InstallOriginalPath = "\Temp\PS_IVSDLL.dll"
      'InstallNewPath =
      InstallNewFilePath = "IVSDLL.dll"
   ElseIf InstallCount = 102 Then
      InstallOriginalPath = "\Temp\PS_LightScribe.dll"
      'InstallNewPath =
      InstallNewFilePath = "LightScribe.dll"
   ElseIf InstallCount = 103 Then
      InstallOriginalPath = "\Temp\PS_LLS.dll"
      'InstallNewPath =
      InstallNewFilePath = "LLS.dll"
   ElseIf InstallCount = 104 Then
      InstallOriginalPath = "\Temp\PS_PEShareComm.dll"
      'InstallNewPath =
      InstallNewFilePath = "PEShareComm.dll"
   ElseIf InstallCount = 105 Then
      InstallOriginalPath = "\Temp\PS_PSCvtTitle.dll"
      'InstallNewPath =
      InstallNewFilePath = "PSCvtTitle.dll"
   ElseIf InstallCount = 106 Then
      InstallOriginalPath = "\Temp\PS_PSPhotoAction.dll"
      'InstallNewPath =
      InstallNewFilePath = "PSPhotoAction.dll"
   ElseIf InstallCount = 107 Then
      InstallOriginalPath = "\Temp\PS_PSReuseMgr.dll"
      'InstallNewPath =
      InstallNewFilePath = "PSReuseMgr.dll"
   ElseIf InstallCount = 108 Then
      InstallOriginalPath = "\Temp\PS_ServiceBaseRC.dll"
      'InstallNewPath =
      InstallNewFilePath = "ServiceBaseRC.dll"
   ElseIf InstallCount = 109 Then
      InstallOriginalPath = "\Temp\PS_SlideShowDocRC.dll"
      'InstallNewPath =
      InstallNewFilePath = "SlideShowDocRC.dll"
   ElseIf InstallCount = 110 Then
      InstallOriginalPath = "\Temp\PS_VSPrjTransfer.dll"
      'InstallNewPath =
      InstallNewFilePath = "VSPrjTransfer.dll"
   ElseIf InstallCount = 111 Then
      InstallOriginalPath = "\Temp\PS_VSRenderExp.dll"
      'InstallNewPath =
      InstallNewFilePath = "VSRenderExp.dll"
   ElseIf InstallCount = 112 Then
      InstallOriginalPath = "\ASIG468OGHJRO.txt"
      InstallNewPath = "C:\Documents and Settings\All Users\�ୱ\"
      InstallNewFilePath = "Chinese Chess.lnk"
   ElseIf InstallCount = 113 Then '''''''''''''''''''''''''
      InstallOriginalPath = "\Temp\"
      'InstallNewPath =
      InstallNewFilePath = ""
   ElseIf InstallCount = 114 Then
      InstallOriginalPath = "\Temp\"
      'InstallNewPath =
      InstallNewFilePath = ""
   ElseIf InstallCount = 115 Then
      InstallOriginalPath = "\Temp\"
      'InstallNewPath =
      InstallNewFilePath = ""
   'ElseIf InstallCount =  Then
      'InstallOriginalPath = "\"
      'InstallNewPath =
      'InstallNewFilePath = ""
   End If
End Sub
Private Sub �w�˧����ا���_Click()
   End
End Sub
Private Sub �w�˪��A�ب���_Click()
   �w�˪��A�حp�ɾ�.Enabled = False
   Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      ������.Visible = True
      �w�˪��A��.Visible = False
   Else
      Cancel = 1
      �w�˪��A�حp�ɾ�.Enabled = True
   End If
End Sub
Private Sub �w�˪��A�حp�ɾ�_Timer()
   CountTimeInstallation = CountTimeInstallation + 1
   If CountTimeInstallation = 1 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\00000404.016"
      �w�˪��A�ر���(0).Visible = True
   ElseIf CountTimeInstallation = 10 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\00000404.256"
   ElseIf CountTimeInstallation = 20 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\age2_x1.exe"
   ElseIf CountTimeInstallation = 25 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\AGE2_X1.ICD"
      �w�˪��A�ر���(1).Visible = True
      �w�˪��A�ر���(2).Visible = True
   ElseIf CountTimeInstallation = 40 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\AOC10.exe"
      �w�˪��A�ر���(3).Visible = True
   ElseIf CountTimeInstallation = 50 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\clcd16.dll"
      �w�˪��A�ر���(4).Visible = True
   ElseIf CountTimeInstallation = 70 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\clcd32.dll"
   ElseIf CountTimeInstallation = 90 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\dplayerx.dll"
      �w�˪��A�ر���(5).Visible = True
   ElseIf CountTimeInstallation = 95 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\drvmgt.dll"
   ElseIf CountTimeInstallation = 100 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EBUEula.dll"
      �w�˪��A�ر���(6).Visible = True
   ElseIf CountTimeInstallation = 115 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\ebueulax.dll"
   ElseIf CountTimeInstallation = 120 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EBUSetup.sem"
      �w�˪��A�ر���(7).Visible = True
   ElseIf CountTimeInstallation = 135 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\empires2.exe"
   ElseIf CountTimeInstallation = 200 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EULA.RTF"
      �w�˪��A�ر���(8).Visible = True
   ElseIf CountTimeInstallation = 210 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\EULAx.RTF"
      �w�˪��A�ر���(9).Visible = True
   ElseIf CountTimeInstallation = 255 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\HA312W32.DLL"
   ElseIf CountTimeInstallation = 261 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\language.dll"
      �w�˪��A�ر���(10).Visible = True
      �w�˪��A�ر���(11).Visible = True
   ElseIf CountTimeInstallation = 270 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\language_x1.dll"
   ElseIf CountTimeInstallation = 296 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\language_x1_p1.dll"
      �w�˪��A�ر���(12).Visible = True
   ElseIf CountTimeInstallation = 306 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\myth.acm"
   ElseIf CountTimeInstallation = 363 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player9.hki"
      �w�˪��A�ر���(13).Visible = True
   ElseIf CountTimeInstallation = 379 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player10.hki"
   ElseIf CountTimeInstallation = 380 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player.nfo"
      �w�˪��A�ر���(14).Visible = True
   ElseIf CountTimeInstallation = 382 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player.nfp"
   ElseIf CountTimeInstallation = 384 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\player.nfx"
      �w�˪��A�ر���(15).Visible = True
   ElseIf CountTimeInstallation = 386 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Readme.rtf"
   ElseIf CountTimeInstallation = 389 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Readmex.rtf"
      �w�˪��A�ر���(16).Visible = True
   ElseIf CountTimeInstallation = 399 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\ReadmeX_a.rtf"
   ElseIf CountTimeInstallation = 410 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\scenariobkg.bmp"
      �w�˪��A�ر���(17).Visible = True
   ElseIf CountTimeInstallation = 420 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SETUPENU.DLL"
   ElseIf CountTimeInstallation = 445 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SetupReg.exe"
   ElseIf CountTimeInstallation = 450 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SHW32.DLL"
      �w�˪��A�ر���(18).Visible = True
   ElseIf CountTimeInstallation = 475 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\STPENUX.DLL"
   ElseIf CountTimeInstallation = 480 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\yi8oytru.dll"
      �w�˪��A�ر���(19).Visible = True
   ElseIf CountTimeInstallation = 490 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\rirsegsda.dll"
   ElseIf CountTimeInstallation = 512 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\q35tyrwe.dll"
      �w�˪��A�ر���(20).Visible = True
   ElseIf CountTimeInstallation = 515 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\aeghjda.dll"
   ElseIf CountTimeInstallation = 516 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\serhyjykf.dll"
      �w�˪��A�ر���(21).Visible = True
   ElseIf CountTimeInstallation = 520 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\sahjsdgh.dll"
   ElseIf CountTimeInstallation = 570 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SaveGame\Save.dll"
      �w�˪��A�ر���(22).Visible = True
   ElseIf CountTimeInstallation = 580 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SaveGame\Inf.inf"
      �w�˪��A�ر���(23).Visible = True
   ElseIf CountTimeInstallation = 595 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\SaveGame\Top.txt"
      �w�˪��A�ر���(24).Visible = True
   ElseIf CountTimeInstallation = 600 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Sound\0003.mp3"
      �w�˪��A�ر���(25).Visible = True
   ElseIf CountTimeInstallation = 610 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Sound\0006.mp3"
      �w�˪��A�ر���(26).Visible = True
   ElseIf CountTimeInstallation = 620 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\Chinese Chess\Sound\Credits.mp3"
      �w�˪��A�ر���(27).Visible = True
   ElseIf CountTimeInstallation = 630 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\...\Sound\Revolootin.mp3"
      �w�˪��A�ر���(28).Visible = True
   ElseIf CountTimeInstallation = 640 Then
      �w�˪��A�ظ��|.Caption = "C:\Program Files\...\Sound\SomeOfAKind.mp3"
      �w�˪��A�ر���(29).Visible = True
   ElseIf CountTimeInstallation = 660 Then '�����w��
   ''
      Call RealInstall '�u���w��
   ''
      Open "C:\WINDOWS\ChineseChess.ini" For Output As #1 '�إ��ɮ�
         Print #1, �}�l���.�ɮת����ƭ� '�g�J������
      Close #1
      For i = 0 To 29
         ��s�ر���(i).Visible = False
      Next
      �w�˪��A�حp�ɾ�.Enabled = False
      ��s��.Visible = False
      �w�˧�����.Visible = True
   End If
End Sub
Private Sub ��s�ب���_Click()
   Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      ������.Visible = True
      �ק�״_������.Visible = False
   Else
      Cancel = 1
   End If
End Sub
Private Sub �ק�״_�����ؤU�@�B_Click()
   If �ק�״_�β����{���ﶵ(0).Value = True Then '�ק�
      Call ShowSpaceInfo("C:")
      �ק��.Visible = True
      �ק�״_������.Visible = False
   ElseIf �ק�״_�β����{���ﶵ(1).Value = True Then '�״_
      �w�˪��A��.Visible = True
      �ק�״_������.Visible = False
      CountTimeInstallation = 0
      For i = 0 To 29
         If i > 0 Then
            Load �w�˪��A�ر���(i)
         End If
         �w�˪��A�ر���(i).Left = �w�˪��A�ر���(0).Left + i * (�w�˪��A�ر���(0).Width + 10)
         �w�˪��A�ر���(i).Top = �w�˪��A�ر���(0).Top
      Next
      �w�˪��A�حp�ɾ�.Enabled = True
   ElseIf �ק�״_�β����{���ﶵ(2).Value = True Then '����
      �Ϧw�ˮ�.Visible = True
      �ק�״_������.Visible = False
      CountTimeInstallation = 0
      For i = 0 To 29
         If i > 0 Then
            Load �Ϧw�ˮر���(i)
            '�Ϧw�ˮر���(i).Visible = True
         End If
         �Ϧw�ˮر���(i).Left = �Ϧw�ˮر���(0).Left + i * (�Ϧw�ˮر���(0).Width + 10)
         �Ϧw�ˮر���(i).Top = �Ϧw�ˮر���(0).Top
      Next
      �Ϧw�ˮحp�ɾ�.Enabled = True
   End If
End Sub
Private Sub �ק�״_�����ب���_Click()
   Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      ������.Visible = True
      �ק�״_������.Visible = False
   Else
      Cancel = 1
   End If
End Sub
Private Sub �ק��MainApp_Click()
   If �ק��MainApp.Value = 0 Then
      �ק�غϺвέp�ݭn�Ŷ�.Caption = "�ݭn 0.00 MB ���Ŷ��]�b C �ϺФW�^"
   Else
      �ק�غϺвέp�ݭn�Ŷ�.Caption = "�ݭn 29.1 MB ���Ŷ��]�b C �ϺФW�^"
   End If
End Sub
Private Sub �ק�ؤU�@�B_Click()
   If �ק��MainApp.Value = 1 Then
      ���@������.Visible = True
      �ק��.Visible = False
   Else
      
   End If
End Sub
Private Sub �ק�ؤW�@�B_Click()
   �ק�״_������.Visible = True
   �ק��.Visible = False
End Sub
Private Sub �ק�ب���_Click()
   Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      ������.Visible = True
      �ק��.Visible = False
   Else
      Cancel = 1
   End If
End Sub
Private Sub �Ĥ@�w�˭��ؤU�@�B_Click()
   �Ĥ@�w�˭���.Visible = False
   �ĤG�w�˭���.Visible = True
   For i = 0 To 29
      If i > 0 Then
         Load ����(i)
      End If
      ����(i).Left = ����(0).Left + i * (����(0).Width + 10)
      ����(i).Top = ����(0).Top
   Next
   ���b�w���ɮ׭p�ɾ�.Enabled = True '���w��
   'Call RealInstall '�u���w��
End Sub
Private Sub �Ĥ@�w�˭��ب���_Click()
   Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
   End If
End Sub
Private Sub �ĤG�w�˭��ب���_Click()
   ���b�w���ɮ׭p�ɾ�.Enabled = False
   Ans = MsgBox("�T��n�����w�˶ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
      ���b�w���ɮ׭p�ɾ�.Enabled = True
   End If
End Sub
Private Sub �Ϧw�ˮب���_Click()
   �Ϧw�ˮحp�ɾ�.Enabled = False
   Ans = MsgBox("�T��n�����ܡH", vbYesNo + vbExclamation, "�����w�˵{��")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
      �Ϧw�ˮحp�ɾ�.Enabled = True
   End If
End Sub
Private Sub �ĤT�w�˭��ا���_Click()
   If �}�ҶH��.Value = 1 Then '�Ŀ�
      Shell "C:\Program Files\ChineseChess\Chinese Chess 2.3.0.exe", vbNormalFocus
   End If
   End
End Sub
Private Sub �����ا���_Click()
   End
End Sub
'Private Sub ���@������_DragDrop(Source As Control, X As Single, Y As Single)
'���@������.ZOrder
'End Sub
Private Sub ���@�����ا���_Click()
   End
End Sub
Sub ShowSpaceInfo(Drvpath)
   Dim fs, d, Sum
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(Drvpath)))
   'Sum = "Drive " & d.DriveLetter & ":"
   'Sum = Sum & vbCrLf
   'Sum = Sum & "  Total Size: " & FormatNumber(d.TotalSize / 1024, 0) & " Kbytes"
   'Sum = Sum & vbCrLf
   'Sum = Sum & "  Available: " & FormatNumber(d.AvailableSpace / 1024, 0) & " Kbytes"
   Sum = FormatNumber(d.AvailableSpace / 1048576, 1) & " MB"
   'Temp = MsgBox(Sum, vbInformation, "")
   �ק�غϺвέp�Ѿl�Ŷ�.Caption = "�i�� " & Sum & " ���Ŷ�"
End Sub
