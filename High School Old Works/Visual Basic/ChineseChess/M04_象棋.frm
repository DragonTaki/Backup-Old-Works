VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form �D�� 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H��"
   ClientHeight    =   12270
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   17925
   Icon            =   "M04_�H��.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12270
   ScaleWidth      =   17925
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame ����� 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  '�S���ؽu
      Height          =   10380
      Left            =   0
      TabIndex        =   28
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.TextBox ��� 
         Appearance      =   0  '����
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   1178
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   41
         ToolTipText     =   "����..."
         Top             =   3360
         Width           =   10695
      End
      Begin VB.TextBox ���� 
         Appearance      =   0  '����
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   1178
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   40
         ToolTipText     =   "����..."
         Top             =   960
         Width           =   10695
      End
      Begin VB.TextBox �s�� 
         Appearance      =   0  '����
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   1178
         Locked          =   -1  'True
         MousePointer    =   14  '�b���P�ݸ��Ϊ�
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "M04_�H��.frx":324A
         ToolTipText     =   "����..."
         Top             =   4200
         Width           =   10695
      End
      Begin VB.CommandButton �T�w 
         BackColor       =   &H006957BD&
         Caption         =   "�T�@�@�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1178
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   30
         ToolTipText     =   "�T�w"
         Top             =   9120
         Width           =   10695
      End
      Begin VB.TextBox �����r 
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   7815
         Left            =   1178
         Locked          =   -1  'True
         MousePointer    =   1  '�b���Ϊ�
         MultiLine       =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "M04_�H��.frx":327F
         ToolTipText     =   "����..."
         Top             =   960
         Width           =   10695
      End
   End
   Begin VB.Frame ������ 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  '�S���ؽu
      Height          =   10380
      Left            =   0
      TabIndex        =   32
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.CommandButton �����T�w 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�T�@�@�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   20.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   33
         ToolTipText     =   "�T�w"
         Top             =   8640
         Width           =   5535
      End
      Begin VB.Label �H�ѳW�h 
         Caption         =   $"M04_�H��.frx":334E
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   15.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   638
         TabIndex        =   35
         ToolTipText     =   "�W�h"
         Top             =   1800
         Width           =   11775
      End
      Begin VB.Label �������D 
         BackColor       =   &H00C0E0FF&
         Caption         =   "�@�@�@�C���W�h"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   2978
         TabIndex        =   34
         Top             =   720
         Width           =   7095
      End
   End
   Begin VB.Frame �D���� 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  '�S���ؽu
      Height          =   10380
      Left            =   0
      TabIndex        =   5
      Top             =   -120
      Width           =   13140
      Begin VB.CommandButton �i�J�C�� 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�i�J�C��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   0
         ToolTipText     =   "�C���}�l"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton �C���]�w 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�C���]�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   1
         ToolTipText     =   "�]�w���֣��W�h��"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton ������s 
         BackColor       =   &H00C0FFFF&
         Caption         =   "���@�@��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   3
         ToolTipText     =   "���󥻵{��"
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton �C���������s 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�C������"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   2
         ToolTipText     =   "�����C���W�h"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton �������s 
         BackColor       =   &H00C0FFFF&
         Caption         =   "���@�@��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         ToolTipText     =   "�����C��"
         Top             =   8160
         Width           =   1935
      End
      Begin VB.Image �D���Ϥ� 
         Height          =   9330
         Left            =   1080
         Picture         =   "M04_�H��.frx":35D3
         Stretch         =   -1  'True
         ToolTipText     =   "����H��"
         Top             =   480
         Width           =   6360
      End
   End
   Begin VB.Frame �]�w�� 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '�S���ؽu
      Height          =   10395
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.CommandButton �^�D�� 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�^�D��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   8160
         Width           =   6135
      End
      Begin VB.Frame ���� 
         BackColor       =   &H00000000&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   26.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   6600
         TabIndex        =   11
         ToolTipText     =   "���O�_�θ��Y"
         Top             =   2640
         Width           =   6135
         Begin VB.OptionButton ������ 
            BackColor       =   &H00000000&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   26.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   1920
            TabIndex        =   13
            ToolTipText     =   "�]�w������"
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton �����} 
            BackColor       =   &H00000000&
            Caption         =   "�}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   26.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   480
            TabIndex        =   12
            ToolTipText     =   "�]�w�����}"
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame ���� 
         BackColor       =   &H00000000&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   26.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   6600
         TabIndex        =   7
         ToolTipText     =   "�]�w���ֶ}��"
         Top             =   5400
         Width           =   6135
         Begin VB.Timer ���֭p�ɾ� 
            Interval        =   1000
            Left            =   3480
            Top             =   600
         End
         Begin VB.OptionButton ������ 
            BackColor       =   &H00000000&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   24
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   1920
            TabIndex        =   9
            ToolTipText     =   "�]�w������"
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton ���ֶ} 
            BackColor       =   &H00000000&
            Caption         =   "�}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   24
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   480
            TabIndex        =   8
            ToolTipText     =   "�]�w���ֶ}"
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin MCI.MMControl ���� 
            Height          =   495
            Left            =   3120
            TabIndex        =   10
            Top             =   1440
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   873
            _Version        =   393216
            BorderStyle     =   0
            PrevVisible     =   0   'False
            NextVisible     =   0   'False
            PauseVisible    =   0   'False
            BackVisible     =   0   'False
            StepVisible     =   0   'False
            RecordVisible   =   0   'False
            EjectVisible    =   0   'False
            DeviceType      =   ""
            FileName        =   ""
         End
      End
      Begin VB.Image �]�w�� 
         Height          =   6015
         Left            =   360
         Picture         =   "M04_�H��.frx":E83E
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   5985
      End
      Begin VB.Label �]�w 
         Alignment       =   2  '�m�����
         BackColor       =   &H00000000&
         Caption         =   "�]�w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   36
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   1920
         TabIndex        =   15
         Top             =   600
         Width           =   9015
      End
   End
   Begin VB.Frame ���s�� 
      BackColor       =   &H005FA76F&
      BorderStyle     =   0  '�S���ؽu
      Height          =   10380
      Left            =   0
      TabIndex        =   36
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.Timer �����Ϥ��p�ɾ� 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   0
         Top             =   0
      End
      Begin VB.Image �����Ϥ� 
         Height          =   1215
         Left            =   3960
         Picture         =   "M04_�H��.frx":12E17
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label �����T�� 
         Alignment       =   2  '�m�����
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   48
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   3818
         TabIndex        =   38
         Top             =   5400
         Width           =   5415
      End
      Begin VB.Label �����N�䭫�s 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '�z��
         Caption         =   "�����N�䭫�s"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   21.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   4838
         TabIndex        =   37
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame �H�Ѯ� 
      BackColor       =   &H005FA76F&
      BorderStyle     =   0  '�S���ؽu
      Caption         =   "s"
      Height          =   10380
      Left            =   0
      TabIndex        =   16
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.Frame �H�ѹϮ� 
         BorderStyle     =   0  '�S���ؽu
         Height          =   1335
         Left            =   11760
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image �ťչ� 
            Height          =   1215
            Left            =   0
            Picture         =   "M04_�H��.frx":12F7A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   31
            Left            =   0
            Picture         =   "M04_�H��.frx":130DD
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   30
            Left            =   0
            Picture         =   "M04_�H��.frx":1361D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   29
            Left            =   0
            Picture         =   "M04_�H��.frx":13DEC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   28
            Left            =   0
            Picture         =   "M04_�H��.frx":1432C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   27
            Left            =   0
            Picture         =   "M04_�H��.frx":1486C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   26
            Left            =   0
            Picture         =   "M04_�H��.frx":14DAC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   25
            Left            =   0
            Picture         =   "M04_�H��.frx":15336
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   24
            Left            =   0
            Picture         =   "M04_�H��.frx":158C0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   23
            Left            =   0
            Picture         =   "M04_�H��.frx":15E36
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   22
            Left            =   0
            Picture         =   "M04_�H��.frx":163AC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   21
            Left            =   0
            Picture         =   "M04_�H��.frx":1695B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   20
            Left            =   0
            Picture         =   "M04_�H��.frx":16F0A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   19
            Left            =   0
            Picture         =   "M04_�H��.frx":173E6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   18
            Left            =   0
            Picture         =   "M04_�H��.frx":178C2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   17
            Left            =   0
            Picture         =   "M04_�H��.frx":17DDB
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   16
            Left            =   0
            Picture         =   "M04_�H��.frx":182F4
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   15
            Left            =   0
            Picture         =   "M04_�H��.frx":18823
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   14
            Left            =   0
            Picture         =   "M04_�H��.frx":18D1F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   13
            Left            =   0
            Picture         =   "M04_�H��.frx":1921B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   12
            Left            =   0
            Picture         =   "M04_�H��.frx":19717
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   11
            Left            =   0
            Picture         =   "M04_�H��.frx":19C13
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   10
            Left            =   0
            Picture         =   "M04_�H��.frx":1A10F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   9
            Left            =   0
            Picture         =   "M04_�H��.frx":1A666
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   8
            Left            =   0
            Picture         =   "M04_�H��.frx":1ABBD
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   7
            Left            =   0
            Picture         =   "M04_�H��.frx":1B10E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   6
            Left            =   0
            Picture         =   "M04_�H��.frx":1B65F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   5
            Left            =   0
            Picture         =   "M04_�H��.frx":1BAF9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   4
            Left            =   0
            Picture         =   "M04_�H��.frx":1BF93
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   3
            Left            =   0
            Picture         =   "M04_�H��.frx":1C4DB
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   2
            Left            =   0
            Picture         =   "M04_�H��.frx":1CA23
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   1
            Left            =   0
            Picture         =   "M04_�H��.frx":1CE66
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image �H�ѹ� 
            Height          =   1215
            Index           =   0
            Left            =   0
            Picture         =   "M04_�H��.frx":1D2A9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Frame ���� 
         BackColor       =   &H005FA76F&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   1935
         Left            =   6120
         TabIndex        =   20
         ToolTipText     =   "���n���ʪ���"
         Top             =   7440
         Width           =   3255
         Begin VB.Image ���ѹϤ� 
            Height          =   1215
            Left            =   1200
            Picture         =   "M04_�H��.frx":1D807
            Stretch         =   -1  'True
            Tag             =   "-1"
            ToolTipText     =   "���n���ʪ���"
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Frame �Ѿl�Ѽ� 
         BackColor       =   &H005FA76F&
         Caption         =   "�Ѿl�Ѽ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   17
         ToolTipText     =   "����ѤU����"
         Top             =   7440
         Width           =   5055
         Begin VB.Label ����Ѵ� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   855
            Left            =   3720
            TabIndex        =   19
            ToolTipText     =   "���ѳѴѼ�"
            Top             =   720
            Width           =   975
         End
         Begin VB.Label �¤�Ѵ� 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   1440
            TabIndex        =   18
            ToolTipText     =   "�´ѳѴѼ�"
            Top             =   720
            Width           =   975
         End
         Begin VB.Image �� 
            Height          =   1215
            Left            =   2520
            Picture         =   "M04_�H��.frx":1D96A
            Stretch         =   -1  'True
            Tag             =   "-1"
            ToolTipText     =   "����"
            Top             =   480
            Width           =   1080
         End
         Begin VB.Image �N 
            Height          =   1215
            Left            =   240
            Picture         =   "M04_�H��.frx":1DE99
            Stretch         =   -1  'True
            Tag             =   "-1"
            ToolTipText     =   "�´�"
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Label ��^ 
         Alignment       =   2  '�m�����
         BackColor       =   &H005FA76F&
         Caption         =   "��^"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   9720
         TabIndex        =   27
         Top             =   8880
         Width           =   2895
      End
      Begin VB.Label ���v 
         Alignment       =   2  '�m�����
         BackColor       =   &H005FA76F&
         Caption         =   "���v"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   9720
         TabIndex        =   26
         Top             =   8220
         Width           =   2895
      End
      Begin VB.Label ���s�}�l 
         Alignment       =   2  '�m�����
         BackColor       =   &H005FA76F&
         Caption         =   "���s�}�l"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   9720
         TabIndex        =   25
         Top             =   7560
         Width           =   2895
      End
      Begin VB.Image �����a�G�� 
         Height          =   570
         Left            =   7125
         Picture         =   "M04_�H��.frx":1E3F7
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Image �����a�@�� 
         Height          =   570
         Left            =   1125
         Picture         =   "M04_�H��.frx":1F352
         Stretch         =   -1  'True
         Top             =   720
         Width           =   555
      End
      Begin VB.Label ���a�@�m�W 
         Alignment       =   2  '�m�����
         BackColor       =   &H005FA76F&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   27.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3720
         TabIndex        =   23
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label ���a�G 
         Alignment       =   2  '�m�����
         BackColor       =   &H005FA76F&
         Caption         =   "���a�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   27.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7710
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label ���a�@ 
         Alignment       =   2  '�m�����
         BackColor       =   &H005FA76F&
         Caption         =   "���a�@"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   27.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1710
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   8
         X1              =   11805
         X2              =   11805
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   7
         X1              =   10485
         X2              =   10485
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   6
         X1              =   9165
         X2              =   9165
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   5
         X1              =   7845
         X2              =   7845
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         X1              =   6525
         X2              =   6525
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   3
         X1              =   5205
         X2              =   5205
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   3885
         X2              =   3885
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   2565
         X2              =   2565
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ���u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   1200
         X2              =   1200
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line ��u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         X1              =   1245
         X2              =   11805
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line ��u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   3
         X1              =   1245
         X2              =   11805
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line ��u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   1245
         X2              =   11805
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line ��u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   1245
         X2              =   11805
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line ��u 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   1245
         X2              =   11805
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Image �� 
         Height          =   1215
         Index           =   0
         Left            =   1365
         Picture         =   "M04_�H��.frx":202AD
         Stretch         =   -1  'True
         Tag             =   "-1"
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label ���a�G�m�W 
         Alignment       =   2  '�m�����
         BackColor       =   &H005FA76F&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   27.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   9750
         TabIndex        =   24
         Top             =   720
         Width           =   2175
      End
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
Dim Name1P As String
Dim Name2P As String
Dim �O�_���� As Boolean
Dim �O�_���� As Boolean
Dim ���w�� As Boolean
Dim CountTime As Integer
Dim �w���Ĥ@���� As Boolean
Dim �w�}�� As Boolean '���C���
Dim �Ѥw½�}(32) As Boolean '�Φ�l�O
Dim �Ů�(32) As Boolean '�Φ�l�O
Dim �{�b�����a�G As Boolean
Dim ���a�@�O�ª� As Boolean
Dim �Ĥ@���O�ۤv���� As Boolean
Dim �ĤG���O�ۤv���� As Boolean
Dim �O�i���ʪ� As Boolean
Dim ���a�@�ӧQ As Boolean
Dim �´ѳӧQ As Boolean
Dim �ѥk�W������ As Boolean
Dim �w����^ As Boolean
Dim �Ĥ@���� As Integer
Dim �ĤG���� As Integer
Dim �Ĥ@���Ѧ�m As Integer
Dim �ĤG���Ѧ�m As Integer
Dim �������� As Integer
Dim �����Ϥ����ʶZ��X As Integer '���ʤ�V�ܼ�
Dim �����Ϥ����ʶZ��Y As Integer
Dim �����T�����ʶZ��X As Integer
Dim �����T�����ʶZ��Y As Integer
Dim Temp1 As Integer
Dim Temp2 As String
Dim ���֭p�ɾ�CountTime As Integer
Private Sub Form_Load()
   ����.Text = "�H�� " & �}�l���.�ɮת���
   ���.Text = "����G2012.12.11."
   �D��.Height = 10780
   �D��.Width = 13140
   �D����.Left = 0
   �D����.Top = 0
   �D����.Height = 10380
   �D����.Width = 18015
   �H�Ѯ�.Left = 0
   �H�Ѯ�.Top = 0
   �H�Ѯ�.Height = 10380
   �H�Ѯ�.Width = 18015
   ���s��.Left = 0
   ���s��.Top = 0
   ���s��.Height = 10380
   ���s��.Width = 18015
   ������.Left = 0
   ������.Top = 0
   ������.Height = 10380
   ������.Width = 18015
   �����.Left = 0
   �����.Top = 0
   �����.Height = 10380
   �����.Width = 18015
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim TheFile As String
   Dim Results As String
   TheFile = "C:\WINDOWS\ChineseChess.ini"
   Results = Dir$(TheFile)
   If Results = "" Then '�ɮפ��s�b
      FileCopy (App.Path & "\�H��" & �}�l���.�ɮת��� & "�w����.exe"), ("C:\WINDOWS\�H��" & �}�l���.�ɮת��� & ".exe")
      FileCopy (App.Path & "\�ť�.che"), ("C:\WINDOWS\�ť�.che")
      FileCopy (App.Path & "\0003.mp3"), ("C:\WINDOWS\0003.mp3")
      FileCopy (App.Path & "\0006.mp3"), ("C:\WINDOWS\0006.mp3")
      FileCopy (App.Path & "\Credits.mp3"), ("C:\WINDOWS\Credits.mp3")
      FileCopy (App.Path & "\Revolootin.mp3"), ("C:\WINDOWS\Revolootin.mp3")
      FileCopy (App.Path & "\SomeOfAKind.mp3"), ("C:\WINDOWS\SomeOfAKind.mp3")
      For i = 0 To 31
         FileCopy (App.Path & "\" & i & ".che"), ("C:\WINDOWS\" & i & ".che")
      Next
      ���w�� = True
      �}�l�p�ɾ�.Enabled = False
      �w�˭�.Show
      Unload �}�l���
   End If
   'SetAttr "C:\WINDOWS\ChineseChess.ini", vbHidden '����
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ���a�@�m�W.Caption = Name1P
   ���a�G�m�W.Caption = Name2P
   Randomize
   '�ʹ�
   For i = 0 To 31
      If i > 0 Then
         Load ��(i)
      End If
      ��(i).Left = ��(0).Left + (i Mod 8) * (��(0).Width + 240) '�H0����m���ǡA��8���l��(�N��n�k���X�檺�e��+���_�j�p)
      ��(i).Top = ��(0).Top + Int(i / 8) * (��(0).Height + 105) '�H0����m���ǡA���㰣8(�N��n���U���X�檺����+���_�j�p)
      ��(i).Visible = True
   Next
   Call �~�P
   �����Ϥ����ʶZ��X = 150 '��w�w���ʪ��t�צb���M�w�A��bX�N���k�A�Y�����A�N��}�l���k�A�Y���t�A�N��_�ʮɩ����A��ȶV�j�A�N���ʶV��
   �����Ϥ����ʶZ��Y = -100 '��bY�N��W�U�A�Y�����A�N��}�l���U�A�Y���t�A�N��_�ʮɩ��W�A��ȶV�j�A�N���ʶV��
   �����T�����ʶZ��X = 150
   �����T�����ʶZ��Y = 100
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   �]�w��.Left = 0
   �]�w��.Top = 0
   Randomize
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   �O�_���� = True
   �O�_���� = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MsgBox ("�C�������I")
   End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�D��
Private Sub ����_Click()
   MsgBox ("�C�������I")
   End
End Sub

Private Sub �C������_Click()
   ������.Visible = True
   �D����.Visible = False
End Sub
Private Sub ����_Click()
   �����.Visible = True
   �D����.Visible = False
   Call ���ë��s
End Sub
Private Sub �i�J�C��_Click()
   Name1P = InputBox("��J�Ĥ@�쪱�a�m�W", "", "�L�W")
   If Name1P <> "" Then
      Name2P = InputBox("��J�ĤG�쪱�a�m�W", "", "�L�W")
      If Name2P <> "" Then
         �H�Ѯ�.Visible = True
         �D����.Visible = False
         Call ���ë��s
      End If
   End If
End Sub
Private Sub �C���]�w_Click()
   Call ���ë��s
   �]�w��.Visible = True
   �D����.Visible = False
End Sub
Private Sub �C���������s_Click()
   Call ���ë��s
   Call �C������_Click
End Sub
Private Sub ������s_Click()
   Call ���ë��s
   Call ����_Click
End Sub
Private Sub �������s_Click()
    Call ����_Click
End Sub
Private Sub ���ë��s()
   ����.Visible = False
   �ɮ�.Visible = False
End Sub
Private Sub ��ܫ��s()
   ����.Visible = True
   �ɮ�.Visible = True
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�H��
Private Sub ��^_Click()
   Ans = MsgBox("�A�T�w�n�����H", vbYesNo, "")
   If Ans = vbYes Then
      Name1P = Empty
      Name2P = Empty
      �w���Ĥ@���� = False
      �w�}�� = False
      For i = 0 To 32
         �Ѥw½�}(i) = False
      Next
      For j = 0 To 32
         �Ů�(j) = False
      Next
      �{�b�����a�G = False
      ���a�@�O�ª� = False
      �Ĥ@���O�ۤv���� = False
      �ĤG���O�ۤv���� = False
      �O�i���ʪ� = False
      ���a�@�ӧQ = False
      �´ѳӧQ = False
      �Ĥ@���� = Empty
      �ĤG���� = Empty
      �Ĥ@���Ѧ�m = Empty
      �ĤG���Ѧ�m = Empty
      �������� = Empty
      �ѥk�W������ = False
      �w����^ = True
      �D����.Visible = True
      �H�Ѯ�.Visible = False
      Call ��ܫ��s
   Else
      Cancel = 1
   End If
End Sub
Private Sub ���s��_Click()
   �����Ϥ��p�ɾ�.Enabled = False
   �H�Ѯ�.Visible = True
   ���s��.Visible = False
   Call ���s�}�l_Click
End Sub
Private Sub �����Ϥ��p�ɾ�_Timer()
   �����Ϥ�.Left = �����Ϥ�.Left + �����Ϥ����ʶZ��X '����X�y�СA����X�Цb<Left>�o�ݩʤ��A��ȶV�j�N��V�k��
   �����Ϥ�.Top = �����Ϥ�.Top + �����Ϥ����ʶZ��Y '����Y�y�СA����X�Цb<Top>�o�ݩʤ��A��ȶV�j�A�N��V�U��
   If �����Ϥ�.Left + �����Ϥ�.Width > �D��.Width Then '�ˬd�O�_�W�L���k��(Form1.Width�N�O���e��)�A���@�Ӫ��󪺥��ɭȬO<Left>�A�k�ɭȬO<Left>+<Width>
      �����Ϥ����ʶZ��X = 0 - Abs(�����Ϥ����ʶZ��X) '�p�G�W�L���ʤ�V�令�t�ȡA�N���V����
   End If
   If �����Ϥ�.Top + �����Ϥ�.Height > �D��.Height - 600 Then '�ˬd�O�_�W�L���U��(Form1.Height�N�O��氪�סA�N�O�t�Ŧ���D�ϡA�ҥH�n��600�A�ϱo��u���U�ɭ�)�A
      �����Ϥ����ʶZ��Y = 0 - Abs(�����Ϥ����ʶZ��Y) '���@�Ӫ��󪺤W�ɭȬO<Top>�A�U�ɭȬO<top>+<Height>,�W�L�N�Ⲿ�ʭȧ令�t���A�N���W��
   End If 'Abs()���@�Ӽƪ����
   If �����Ϥ�.Left < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G���ɭ�<Left>�p��0�A�N���쥪��F�A�N��ȴ��������A�啕�k��
      �����Ϥ����ʶZ��X = Abs(�����Ϥ����ʶZ��X)
   End If
   If �����Ϥ�.Top < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G�W�ɭ�<Top>�p��0�A�N����W��F�A�N��ȴ��������A�啕�U��
      �����Ϥ����ʶZ��Y = Abs(�����Ϥ����ʶZ��Y)
   End If
   ''''''''''''''''''''''''''''''''''''''''
   �����T��.Left = �����T��.Left + �����T�����ʶZ��X '����X�y�СA����X�Цb<Left>�o�ݩʤ��A��ȶV�j�N��V�k��
   �����T��.Top = �����T��.Top + �����T�����ʶZ��Y '����Y�y�СA����X�Цb<Top>�o�ݩʤ��A��ȶV�j�A�N��V�U��
   If �����T��.Left + �����T��.Width > �D��.Width Then '�ˬd�O�_�W�L���k��(Form1.Width�N�O���e��)�A���@�Ӫ��󪺥��ɭȬO<Left>�A�k�ɭȬO<Left>+<Width>
      �����T�����ʶZ��X = 0 - Abs(�����T�����ʶZ��X) '�p�G�W�L���ʤ�V�令�t�ȡA�N���V����
   End If
   If �����T��.Top + �����T��.Height > �D��.Height - 600 Then '�ˬd�O�_�W�L���U��(Form1.Height�N�O��氪�סA�N�O�t�Ŧ���D�ϡA�ҥH�n��600�A�ϱo��u���U�ɭ�)�A
      �����T�����ʶZ��Y = 0 - Abs(�����T�����ʶZ��Y) '���@�Ӫ��󪺤W�ɭȬO<Top>�A�U�ɭȬO<top>+<Height>,�W�L�N�Ⲿ�ʭȧ令�t���A�N���W��
   End If  'Abs()���@�Ӽƪ����
   If �����T��.Left < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G���ɭ�<Left>�p��0�A�N���쥪��F�A�N��ȴ��������A�啕�k��
      �����T�����ʶZ��X = Abs(�����T�����ʶZ��X)
   End If
   If �����T��.Top < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G�W�ɭ�<Top>�p��0�A�N����W��F�A�N��ȴ��������A�啕�U��
      �����T�����ʶZ��Y = Abs(�����T�����ʶZ��Y)
   End If
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
Private Sub ��_Click(Index As Integer)
   If �w�}�� = False Then '�M�w�C��
      If 16 > ��(Index).Tag And ��(Index).Tag >= 0 Then '�I��´�
         ���a�@.ForeColor = vbBlack
         ���a�G.ForeColor = vbRed
         ���a�@�O�ª� = True
      Else '�I�����
         ���a�@.ForeColor = vbRed
         ���a�G.ForeColor = vbBlack
      End If
      Call ���H
      �w�}�� = True
      ��(Index).Picture = �H�ѹ�(��(Index).Tag).Picture
      �Ѥw½�}(Index) = True
   Else '�w�M�w�C��
      If �Ѥw½�}(Index) = False Then '�Ӵѥ�½�}
         ��(Index).Picture = �H�ѹ�(��(Index).Tag).Picture
         �Ѥw½�}(Index) = True
         Call ���H
      ElseIf �w���Ĥ@���� = True And �Ѥw½�}(Index) = False Then '�ĤG���ѥ�½�}
         If �O�_�t�Y = True Then '�i�t�Y
            GoTo �i�t�Y '���P�w½�}
         Else '���i�t�Y�A���}
            Exit Sub
         End If
         '�٨S��
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
            ���ѹϤ�.Picture = �H�ѹ�(��(Index).Tag).Picture
         End If
      ElseIf �w���Ĥ@���� = True And �Ѥw½�}(Index) = True Then '�ӴѤw½�}�A�O�ĤG��
         If ��(Index).Tag = �Ĥ@���� Then '�M�Ĥ@���ѬۦP�A��������
            �w���Ĥ@���� = False '�٨S��
            �Ĥ@���� = Empty
            �Ĥ@���Ѧ�m = Empty
            ���ѹϤ�.Picture = �ťչ�.Picture
         End If
         'If �Ĥ@���� = 9 Or �Ĥ@���� = 10 Then '�ҥ~�G�Ĥ@���ѬO�]
         '   If 32 > ��(Index).Tag And ��(Index).Tag >= 27 Then '�ĤG���ѬO�L
         '       Exit Sub '���P����
         '   End If
         'ElseIf �Ĥ@���� = 25 Or �Ĥ@���� = 26 Then '�ҥ~�G�Ĥ@���ѬO��
         '   If 16 > ��(Index).Tag And ��(Index).Tag >= 15 Then '�ĤG���ѬO��
         '       Exit Sub '���P����
         '   End If
         'End If
�i�t�Y:
         �ĤG���� = ��(Index).Tag
         �ĤG���Ѧ�m = Index
         Call �P�_�ĤG���O�_���ۤv����
         If �ĤG���O�ۤv���� = True Then
            �ĤG���O�ۤv���� = False
            �ĤG���� = Empty
            �ĤG���Ѧ�m = Empty
            Exit Sub '�ĤG���O�ۤv���ѡA����Y�A���}
         Else '�ĤG�����O�ۤv���ѡA�i�H�Y
            Call �ˬd�����Ѫ���m
            If �O�i���ʪ� = False Then
               Exit Sub
            End If
            �O�i���ʪ� = False
            �������� = 0
            Call �ʴѤl
         End If
      End If
   End If
   If Val(����Ѵ�.Caption) = 0 Then '�����C���A�´ѳӧQ
      If ���a�@�O�ª� = True Then '���a�@�ӧQ
         ���a�@�ӧQ = True
         �´ѳӧQ = True
         Call �����C��
      Else '���a�G�ӧQ
         ���a�@�ӧQ = False
         �´ѳӧQ = True
         Call �����C��
      End If
   ElseIf Val(�¤�Ѵ�.Caption) = 0 Then '�����C���A���ѳӧQ
      If ���a�@�O�ª� = False Then '���a�@�ӧQ
         ���a�@�ӧQ = True
         �´ѳӧQ = False
         Call �����C��
      Else '���a�G�ӧQ
         ���a�@�ӧQ = False
         �´ѳӧQ = False
         Call �����C��
      End If
   End If
End Sub
Private Sub �P�_�Ĥ@���O�_���ۤv����()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If 16 > �Ĥ@���� And �Ĥ@���� >= 0 Then '�Ĥ@���O�´�
            �Ĥ@���O�ۤv���� = True
         End If
      Else '���a�@�O����
         If 32 > �Ĥ@���� And �Ĥ@���� >= 16 Then '�Ĥ@���O����
            �Ĥ@���O�ۤv���� = True
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If 32 > �Ĥ@���� And �Ĥ@���� >= 16 Then '�Ĥ@���O����
            �Ĥ@���O�ۤv���� = True
         End If
      Else '���a�G�O�ª�
         If 16 > �Ĥ@���� And �Ĥ@���� >= 0 Then '�Ĥ@���O�´�
            �Ĥ@���O�ۤv���� = True
         End If
      End If
   End If
End Sub
Private Sub �P�_�ĤG���O�_���ۤv����()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If 16 > �ĤG���� And �ĤG���� >= 0 Then '�ĤG���O�´�
            �ĤG���O�ۤv���� = True
         End If
      Else '���a�@�O����
         If 32 > �ĤG���� And �ĤG���� >= 16 Then '�ĤG���O����
            �ĤG���O�ۤv���� = True
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If 32 > �ĤG���� And �ĤG���� >= 16 Then '�ĤG���O����
            �ĤG���O�ۤv���� = True
         End If
      Else '���a�G�O�ª�
         If 16 > �ĤG���� And �ĤG���� >= 0 Then '�ĤG���O�´�
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
   If �Ĥ@���� = 9 Or �Ĥ@���� = 10 Or �Ĥ@���� = 25 Or �Ĥ@���� = 26 Then  '�Ĥ@���ѬO�]�ά��A�b������
      If Abs((�Ĥ@���Ѧ�m - �ĤG���Ѧ�m)) < 8 Then '�Ĥ@���ѩM�ĤG���ѦP��
         If ��(�ĤG���Ѧ�m).Tag = -1 Then '�ĤG���ѬO�Ū�
            If �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 1 Or �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 1 Then '�ĤG���ѬO�Ū��B�b�Ĥ@���ѤW��ΤU��
               Call �ʬ����M��
            End If
         Else '�ĤG���Ѥ��O�Ū�
            Temp1 = �Ĥ@���Ѧ�m
            Temp2 = �ĤG���Ѧ�m
            If Temp1 > Temp2 Then
               Temp1 = �ĤG���Ѧ�m
               Temp2 = �Ĥ@���Ѧ�m
            End If
            For i = Temp1 + 1 To Temp2
               If ��(i).Tag <> -1 Then
                  �������� = �������� + 1
               End If
            Next
            If �������� = 2 Then
               Call �ʬ����M��
            End If
         End If
      ElseIf Abs(�Ĥ@���Ѧ�m - �ĤG���Ѧ�m) Mod 8 = 0 Then  '�Ĥ@���ѩM�ĤG���ѦP�C
         If ��(�ĤG���Ѧ�m).Tag = -1 Then '�ĤG���ѬO�Ū�
            If �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 8 Or �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 8 Then '
               Call �ʬ����M��
            End If
         Else '�ĤG���Ѥ��O�Ū�
            Temp1 = �Ĥ@���Ѧ�m
            Temp2 = �ĤG���Ѧ�m
            If Temp1 > Temp2 Then
               Temp1 = �ĤG���Ѧ�m
               Temp2 = �Ĥ@���Ѧ�m
            End If
            For i = Temp1 + 8 To Temp2 Step 8
               If ��(i).Tag <> -1 Then
                  �������� = �������� + 1
               End If
            Next
            If �������� = 2 Then
               Call �ʬ����M��
            End If
         End If
      End If
   Else '�Ĥ@���Ѥ��O�]�ά�
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
   End If
End Sub
Private Sub �Y���M��()
   �w���Ĥ@���� = False
   �Ĥ@���� = Empty
   �ĤG���� = Empty
   �Ĥ@���Ѧ�m = Empty
   �ĤG���Ѧ�m = Empty
   ���ѹϤ�.Picture = �ťչ�.Picture
   Call ���H
End Sub
Private Sub �ʬ����M��()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         End If
      Else '���a�@�O����
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         End If
      Else '���a�G�O�ª�
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         End If
      End If
   End If
   ��(�ĤG���Ѧ�m).Picture = �H�ѹ�(�Ĥ@����).Picture
   ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
   ���ѹϤ�.Picture = �ťչ�.Picture
   Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
   �Ů�(�Ĥ@���Ѧ�m) = True
   �Ĥ@���� = Empty
   �ĤG���� = Empty
   �Ĥ@���Ѧ�m = Empty
   �ĤG���Ѧ�m = Empty
   �w���Ĥ@���� = False
   �������� = 0
   Call ���H
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
Private Sub ���Ѥl�@�ܴѤl�G()
   ��(�ĤG���Ѧ�m).Picture = �H�ѹ�(�Ĥ@����).Picture
   ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
End Sub
Private Sub �ʶ´�()
   If �Ĥ@���� = 0 Then '�N
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 27 > �ĤG���� And �ĤG���� >= 16 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 32 > �ĤG���� And �ĤG���� >= 27 Then '�N�Y�L
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 1 Or �Ĥ@���� = 2 Then '�h
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 17 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 16 Then '�h�Y��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 3 Or �Ĥ@���� = 4 Then '�H
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 19 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Then '�H�Y�ӥK
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 5 Or �Ĥ@���� = 6 Then '��
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 21 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 21 > �ĤG���� And �ĤG���� >= 16 Then '���Y�ӥK��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 7 Or �Ĥ@���� = 8 Then '��
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 23 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 23 > �ĤG���� And �ĤG���� >= 16 Then '���Y�ӥK����
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 9 Or �Ĥ@���� = 10 Then '�]
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 27 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 27 > �ĤG���� And �ĤG���� >= 16 Then
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf 16 > �Ĥ@���� And �Ĥ@���� >= 11 Then '��
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 25 Or �ĤG���� = 26 Then '�Y��
         '����
      ElseIf �ĤG���� = 16 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 25 > �ĤG���� And �ĤG���� >= 17 Then    '��Y�K���ϽX
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   End If
End Sub
Private Sub �ʬ���()
   If �Ĥ@���� = 16 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 11 > �ĤG���� And �ĤG���� >= 0 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 15 > �ĤG���� And �ĤG���� > 11 Then '�ӦY��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 17 Or �Ĥ@���� = 18 Then '�K
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then   '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Then '�K�Y�N
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 19 Or �Ĥ@���� = 20 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Then '�ۦY�N�h
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 21 Or �Ĥ@���� = 22 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Then '�ۦY�N�h�H
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 23 Or �Ĥ@���� = 24 Then '�X
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Then '�ۦY�N�h�H��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 25 Or �Ĥ@���� = 26 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Then
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf 32 > �Ĥ@���� And �Ĥ@���� >= 27 Then '�L
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 9 Or �ĤG���� = 10 Then '�Y�]
         '����
      ElseIf �ĤG���� = 0 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".che")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Then '�L�Y�h�H����
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   End If
End Sub
Private Sub �����C��()
   �����Ϥ��p�ɾ�.Enabled = True
   ���s��.Visible = True
   �H�Ѯ�.Visible = False
   If ���a�@�ӧQ = True Then
      �����T��.Caption = ���a�@�m�W & "�ӧQ"
   Else
      �����T��.Caption = ���a�G�m�W & "�ӧQ"
   End If
End Sub
Private Sub ���s�}�l_Click()
   Call �~�P
   For i = 0 To 31 '�\�P
      ��(i).Picture = �ťչ�.Picture
   Next
   �¤�Ѵ�.Caption = "16"
   ����Ѵ�.Caption = "16"
   �����a�@��.Visible = True
   �����a�G��.Visible = False
   �{�b�����a�G = False
   �w���Ĥ@���� = False
   For i = 0 To 32
      �Ѥw½�}(i) = False
   Next
   For j = 0 To 32
      �Ů�(j) = False
   Next
   �w�}�� = False
   �{�b�����a�G = False
   ���a�@�O�ª� = False
   �Ĥ@���O�ۤv���� = False
   �ĤG���O�ۤv���� = False
   �O�i���ʪ� = False
   ���a�@�ӧQ = False
   �´ѳӧQ = False
   �Ĥ@���� = Empty
   �ĤG���� = Empty
   �Ĥ@���Ѧ�m = Empty
   �ĤG���Ѧ�m = Empty
   �������� = Empty
   ���ѹϤ�.Picture = �ťչ�.Picture
End Sub
Private Sub ���v_Click()
   If �{�b�����a�G = False Then '���a�@���v
      ���a�@�ӧQ = False
      Call �����C��
   Else '���a�G���v
      ���a�@�ӧQ = True
      Call �����C��
   End If
   �H�Ѯ�.Visible = False
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�]�w
Private Sub �^�D��_Click()
   �D����.Visible = True
   �]�w��.Visible = False
   Call ��ܫ��s
End Sub
Private Sub �����}_Click()
   �O�_���� = True
End Sub
Private Sub ������_Click()
   �O�_���� = False
End Sub
Private Sub ���֭p�ɾ�_Timer()
   ���֭p�ɾ�CountTime = ���֭p�ɾ�CountTime + 1
   If Temp2 = "0003" Then
      If ���֭p�ɾ�CountTime = 102 Then
         ����.Command = "CLOSE"
         ���֭p�ɾ�CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "0006" Then
      If ���֭p�ɾ�CountTime = 73 Then
         ����.Command = "CLOSE"
         ���֭p�ɾ�CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "Credits" Then
      If ���֭p�ɾ�CountTime = 205 Then
         ����.Command = "CLOSE"
         ���֭p�ɾ�CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "Revolootin" Then
      If ���֭p�ɾ�CountTime = 72 Then
         ����.Command = "CLOSE"
         ���֭p�ɾ�CountTime = 0
         Call ���ּ���
      End If
   ElseIf Temp2 = "SomeOfAKind" Then
      If ���֭p�ɾ�CountTime = 76 Then
         ����.Command = "CLOSE"
         ���֭p�ɾ�CountTime = 0
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
Private Sub �����T�w_Click()
   �D����.Visible = True
   ������.Visible = False
   Call ��ܫ��s
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
Private Sub �s��_Click()
   Shell "C:\Program Files\Internet Explorer\iexplore.exe & http://www.facebook.com/#!/groups/269973856388003/", vbMaximizedFocus
End Sub
Private Sub �s��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   �s��.ForeColor = vbRed
End Sub
Private Sub �s��_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   �s��.ForeColor = &H80FF&
End Sub
Private Sub �T�w_Click()
   �D����.Visible = True
   �����.Visible = False
   Call ��ܫ��s
End Sub
