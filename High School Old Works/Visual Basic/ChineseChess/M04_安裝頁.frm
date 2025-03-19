VERSION 5.00
Begin VB.Form 安裝頁 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋 - InstallShield Wizard"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15075
   Icon            =   "M04_安裝頁.frx":0000
   LinkMode        =   1  '來源
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   15075
   StartUpPosition =   2  '螢幕中央
   Visible         =   0   'False
   Begin VB.Frame 反安裝框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Timer 反安裝框計時器 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7080
         Top             =   0
      End
      Begin VB.Frame 反安裝框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   6120
         TabIndex        =   107
         Top             =   5760
         Width           =   1215
         Begin VB.Label 反安裝框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            Height          =   255
            Left            =   30
            TabIndex        =   108
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 反安裝框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox 反安裝框InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
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
      Begin VB.Label 反安裝框底層背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -120
         TabIndex        =   105
         Top             =   5400
         Width           =   7815
      End
      Begin VB.Shape 反安裝框條型 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   270
         Index           =   0
         Left            =   300
         Shape           =   4  '圓角矩形
         Top             =   2550
         Width           =   135
      End
      Begin VB.Shape 反安裝框條型框 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         Shape           =   4  '圓角矩形
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label 反安裝框移除 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "移除"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 象棋安裝程式正在移除程式 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "象棋 安裝程式正在移除程式。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 移除象棋 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "移除 象棋。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 移除 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "移除"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 反安裝框背景 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame 安裝狀態框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Frame 安裝狀態框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   6120
         TabIndex        =   75
         Top             =   5760
         Width           =   1215
         Begin VB.Label 安裝狀態框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            Height          =   255
            Left            =   30
            TabIndex        =   76
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 安裝狀態框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Timer 安裝狀態框計時器 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7080
         Top             =   0
      End
      Begin VB.TextBox 安裝狀態框InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
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
      Begin VB.Shape 安裝狀態框條型 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   270
         Index           =   0
         Left            =   300
         Shape           =   4  '圓角矩形
         Top             =   2550
         Width           =   135
      End
      Begin VB.Shape 安裝狀態框條型框 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         Shape           =   4  '圓角矩形
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label 安裝狀態 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "安裝狀態"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 象棋安裝程式正在執行所要求的安裝 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "象棋 安裝程式正在執行所要求的安裝。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 安裝狀態框安裝 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "安裝"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 安裝狀態框路徑 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "C:\"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 安裝狀態框背景 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label 安裝狀態框底層背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -240
         TabIndex        =   73
         Top             =   5400
         Width           =   7815
      End
   End
   Begin VB.Frame 修改框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
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
      Begin VB.PictureBox 修改框修改框 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BorderStyle     =   0  '沒有框線
         FillColor       =   &H00FF0000&
         FillStyle       =   0  '實心
         ForeColor       =   &H00FF0000&
         Height          =   2500
         Left            =   250
         ScaleHeight     =   2505
         ScaleWidth      =   3765
         TabIndex        =   54
         Top             =   1790
         Width           =   3760
         Begin VB.CheckBox 修改框MainApp 
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            Caption         =   "Main App"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   46
            Top             =   30
            Value           =   1  '核取
            Width           =   1095
         End
         Begin VB.Line 修改框線 
            BorderStyle     =   3  '點線
            X1              =   120
            X2              =   360
            Y1              =   150
            Y2              =   150
         End
      End
      Begin VB.Frame 修改框說明 
         BackColor       =   &H00E0E0E0&
         Caption         =   "說明"
         ForeColor       =   &H80000002&
         Height          =   2655
         Left            =   4320
         TabIndex        =   53
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox 修改框InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
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
      Begin VB.Frame 修改框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   3480
         TabIndex        =   42
         Top             =   5760
         Width           =   3855
         Begin VB.Label 修改框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            Height          =   255
            Left            =   2670
            TabIndex        =   45
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 修改框下一步 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "下一步(N)　>"
            Height          =   255
            Left            =   1320
            TabIndex        =   44
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 修改框上一步 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "<　上一步(B)"
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 修改框上一步框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 修改框下一步框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 修改框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   2640
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Label 修改框磁碟統計剩餘空間 
         BackColor       =   &H00FFFFFF&
         Caption         =   "可用 0.00 MB 的空間"
         Height          =   255
         Left            =   240
         TabIndex        =   99
         Top             =   4680
         Width           =   3735
      End
      Begin VB.Label 修改框磁碟統計需要空間 
         BackColor       =   &H00FFFFFF&
         Caption         =   "需要 29.1 MB 的空間（在 C 磁碟上）"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   4440
         Width           =   3735
      End
      Begin VB.Shape 修改框元件框 
         BorderColor     =   &H00FF8080&
         FillColor       =   &H00FFC0C0&
         Height          =   2550
         Left            =   240
         Top             =   1770
         Width           =   3795
      End
      Begin VB.Label 修改框底層背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -120
         TabIndex        =   52
         Top             =   5400
         Width           =   7815
      End
      Begin VB.Label 選擇元件 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "選擇元件"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 選擇安裝程式將安裝的元件 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "選擇安裝程式將安裝的元件。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 修改框文字 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "選擇您想要安裝的元件，不選中要卸載的元件。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 修改框背景 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame 修改修復移除框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Frame 修改修復移除框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   5760
         Width           =   3855
         Begin VB.Label 修改修復移除框上一步 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "<　上一步(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 修改修復移除框下一步 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "下一步(N)　>"
            Height          =   255
            Left            =   1320
            TabIndex        =   35
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 修改修復移除框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            Height          =   255
            Left            =   2670
            TabIndex        =   34
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 修改修復移除框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   2640
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 修改修復移除框上一步框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   0
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 修改修復移除框下一步框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox 修改修復移除框InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
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
      Begin VB.OptionButton 修改修復或移除程式選項 
         BackColor       =   &H00E0E0E0&
         Caption         =   "修復(E)"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   3048
         Width           =   975
      End
      Begin VB.OptionButton 修改修復或移除程式選項 
         BackColor       =   &H00E0E0E0&
         Caption         =   "移除(R)"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   11
         Top             =   4176
         Width           =   975
      End
      Begin VB.OptionButton 修改修復或移除程式選項 
         BackColor       =   &H00E0E0E0&
         Caption         =   "修改(M)"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   1920
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label 移除文字 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "移除所有已安裝元件。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 修復文字 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "重新安裝以前安裝程式所安裝的所有程式元件。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 修改文字 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "選擇要新增的新程式元件或選擇要移除的目前已安裝元件。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 修改修復移除框文字 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "歡迎使用 象棋 安裝維護程式。使用此程式可以修改目前的安裝。請按一下下列一個選項。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 修改修復或移除程式 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "修改、修復或移除程式。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 歡迎 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "歡迎"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 修改修復移除框背景 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label 修改修復移除框底層背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -120
         TabIndex        =   16
         Top             =   5400
         Width           =   7815
      End
   End
   Begin VB.Frame 安裝完成框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   -120
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Frame 安裝完成框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   3480
         TabIndex        =   92
         Top             =   5760
         Width           =   3855
         Begin VB.Label 安裝完成框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2670
            TabIndex        =   95
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 安裝完成框完成 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "完成"
            Height          =   255
            Left            =   1395
            TabIndex        =   94
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 安裝完成框完成框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 安裝完成框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   2640
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label 安裝完成框上一步 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "<　上一步(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   0
            TabIndex        =   93
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 安裝完成框上一步框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   0
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox 安裝完成文字 
         Appearance      =   0  '平面
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         Text            =   "M04_安裝頁.frx":324A
         Top             =   1080
         Width           =   3255
      End
      Begin VB.PictureBox 安裝完成框InstallShield圖 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   0
         Picture         =   "M04_安裝頁.frx":3273
         ScaleHeight     =   4155
         ScaleWidth      =   2490
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   0
         Width           =   2490
      End
      Begin VB.TextBox 安裝完成框InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
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
      Begin VB.Label 安裝完成 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "安裝完成"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 安裝完成框背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -1200
         TabIndex        =   96
         Top             =   5400
         Width           =   8775
      End
   End
   Begin VB.Frame 維護完成框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   -120
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox 維護完成框InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
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
      Begin VB.PictureBox 維護完成框InstallShield圖 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   0
         Picture         =   "M04_安裝頁.frx":80CF
         ScaleHeight     =   4155
         ScaleWidth      =   2490
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   0
         Width           =   2490
      End
      Begin VB.TextBox 維護完成文字 
         Appearance      =   0  '平面
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         Text            =   "M04_安裝頁.frx":CF2B
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Frame 維護完成框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   3480
         TabIndex        =   57
         Top             =   5760
         Width           =   3855
         Begin VB.Label 維護完成框上一步 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "<　上一步(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 維護完成框完成 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "完成"
            Height          =   255
            Left            =   1395
            TabIndex        =   59
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 維護完成框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2670
            TabIndex        =   58
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 維護完成框完成框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1320
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 維護完成框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   2640
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 維護完成框上一步框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   0
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Label 維護完成框背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -1200
         TabIndex        =   63
         Top             =   5400
         Width           =   8775
      End
      Begin VB.Label 維護完成 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "維護完成"
         BeginProperty Font 
            Name            =   "新細明體"
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
   Begin VB.Frame 結束框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   -120
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.PictureBox 結束框InstallShield圖 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   0
         Picture         =   "M04_安裝頁.frx":CF5C
         ScaleHeight     =   4155
         ScaleWidth      =   2490
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   0
         Width           =   2490
      End
      Begin VB.Frame 結束框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   3480
         TabIndex        =   37
         Top             =   5760
         Width           =   3855
         Begin VB.Label 結束框上一步 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "<　上一步(B)"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 結束框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2670
            TabIndex        =   40
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label 結束框完成 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "完成"
            Height          =   255
            Left            =   1395
            TabIndex        =   39
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 結束框完成框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   1380
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 結束框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   2640
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
         Begin VB.Shape 結束框上一步框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H8000000B&
            Height          =   300
            Left            =   120
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.TextBox 完成文字 
         Appearance      =   0  '平面
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         Text            =   "M04_安裝頁.frx":11DB8
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label InstallShieldWizard完成 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "InstallShield Wizard 完成"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 結束框背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -1200
         TabIndex        =   20
         Top             =   5400
         Width           =   8775
      End
   End
   Begin VB.Frame 更新框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox 更新框InstallShield 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
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
      Begin VB.Frame 更新框按鈕框 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '沒有框線
         Height          =   375
         Left            =   6120
         TabIndex        =   78
         Top             =   5760
         Width           =   1215
         Begin VB.Label 更新框取消 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
            Caption         =   "取消"
            Height          =   255
            Left            =   30
            TabIndex        =   79
            Top             =   45
            Width           =   1140
         End
         Begin VB.Shape 更新框取消框 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  '不透明
            BorderColor     =   &H00000000&
            Height          =   300
            Left            =   0
            Shape           =   4  '圓角矩形
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Shape 更新框條型框 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         Shape           =   4  '圓角矩形
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Shape 更新框條型 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   270
         Index           =   0
         Left            =   300
         Shape           =   4  '圓角矩形
         Top             =   2550
         Width           =   135
      End
      Begin VB.Label 更新 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "更新"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 更新框底層背景 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   -240
         TabIndex        =   86
         Top             =   5400
         Width           =   7815
      End
      Begin VB.Label 更新框路徑 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "C:\"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 更新框更新 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "更新"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 象棋安裝程式正在更新至最新版本 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "象棋 安裝程式正在更新至最新版本。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 更新象棋 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "更新 象棋。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 更新框背景 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame 第二安裝頁框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   7560
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Timer 正在安裝檔案計時器 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7080
         Top             =   0
      End
      Begin VB.CommandButton 第二安裝頁框取消 
         Caption         =   "取消"
         Height          =   375
         Left            =   4920
         TabIndex        =   24
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton 第二安裝頁框安裝 
         Caption         =   "安裝"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label 第二安裝頁框請稍候 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "請稍候，安裝程式正在將 象棋 安裝到您的電腦上。"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 第二安裝頁框正在安裝 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "正在安裝"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Image 第二安裝頁框圖 
         Height          =   3015
         Left            =   720
         Picture         =   "M04_安裝頁.frx":11E53
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Shape 條型 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   270
         Index           =   0
         Left            =   780
         Shape           =   4  '圓角矩形
         Top             =   2070
         Width           =   135
      End
      Begin VB.Shape 第二安裝頁框條型框 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   720
         Shape           =   4  '圓角矩形
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label 第二安裝頁框正在安裝檔案名稱 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 第二安裝頁框正在安裝檔案 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "正在安裝檔案"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 第二安裝頁框背景 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '單線固定
         Height          =   1095
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.Frame 第一安裝頁框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   7560
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox 版本 
         Appearance      =   0  '平面
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         Text            =   "M04_安裝頁.frx":1D0BE
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton 第一安裝頁框下一步 
         Caption         =   "安裝"
         Height          =   375
         Left            =   3000
         TabIndex        =   0
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton 第一安裝頁框取消 
         Caption         =   "取消"
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox 第一安裝頁框安裝說明 
         Appearance      =   0  '平面
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         Text            =   "M04_安裝頁.frx":1D0CA
         Top             =   480
         Width           =   6735
      End
   End
   Begin VB.Frame 第三安裝頁框 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  '沒有框線
      Height          =   6300
      Left            =   7560
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.CheckBox 開啟象棋 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   2760
         TabIndex        =   32
         Top             =   5160
         Value           =   1  '核取
         Width           =   1695
      End
      Begin VB.CommandButton 第三安裝頁框完成 
         Caption         =   "完成"
         Height          =   375
         Left            =   4920
         TabIndex        =   31
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Image 第三安裝頁框圖 
         Height          =   4290
         Left            =   720
         Picture         =   "M04_安裝頁.frx":1D196
         Stretch         =   -1  'True
         Top             =   360
         Width           =   5250
      End
   End
End
Attribute VB_Name = "安裝頁"
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
   安裝頁.Height = 6700
   安裝頁.Width = 7500
   第一安裝頁框.Left = 0
   第一安裝頁框.Top = 0
   第一安裝頁框.Height = 12840
   第一安裝頁框.Width = 7500
   第二安裝頁框.Left = 0
   第二安裝頁框.Top = 0
   第二安裝頁框.Height = 12840
   第二安裝頁框.Width = 7500
   第三安裝頁框.Left = 0
   第三安裝頁框.Top = 0
   第三安裝頁框.Height = 12840
   第三安裝頁框.Width = 7500
   反安裝框.Left = 0
   反安裝框.Top = 0
   反安裝框.Height = 12840
   反安裝框.Width = 7500
   更新框.Left = 0
   更新框.Top = 0
   更新框.Height = 12840
   更新框.Width = 7500
   結束框.Left = 0
   結束框.Top = 0
   結束框.Height = 12840
   結束框.Width = 7500
   修改框.Left = 0
   修改框.Top = 0
   修改框.Height = 12840
   修改框.Width = 7500
   維護完成框.Left = 0
   維護完成框.Top = 0
   維護完成框.Height = 12840
   維護完成框.Width = 7500
   安裝完成框.Left = 0
   安裝完成框.Top = 0
   安裝完成框.Height = 12840
   安裝完成框.Width = 7500
   安裝狀態框.Left = 0
   安裝狀態框.Top = 0
   安裝狀態框.Height = 12840
   安裝狀態框.Width = 7500
   修改修復移除框.Left = 0
   修改修復移除框.Top = 0
   修改修復移除框.Height = 12840
   修改修復移除框.Width = 7500
   '''''''''''''''''''''''''''''''''''''''''''''''''判斷是否安裝及需更新
   TheFile = "C:\WINDOWS\ChineseChess.ini"
   Results = Dir$(TheFile)
   If Results = "" Then '檔案不存在
      第一安裝頁框.Visible = True
   Else '檔案存在
      Open "C:\WINDOWS\ChineseChess.ini" For Input As #1 '讀檔
         Line Input #1, Temp
      Close #1
      If Val(Temp) < 開始表單.檔案版本數值 Then '較舊版本，需更新
         安裝頁.Width = 7500
         更新框.Visible = True
      ElseIf Val(Temp) = 開始表單.檔案版本數值 Then '版本相同，修改修復移除
         安裝頁.Width = 7500
         修改修復移除框.Visible = True
      Else '較新版本，關閉
         '
         安裝頁.Width = 7500
      End If
   End If
   版本.Text = "象棋 " & 開始表單.檔案版本
   開啟象棋.Caption = "開啟象棋 " & 開始表單.檔案版本
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If 維護完成框.Visible = True Or 結束框.Visible = True Or 安裝完成框.Visible = True Then
      End
   ElseIf 反安裝框.Visible = True Then
         反安裝框計時器.Enabled = False
      Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
      If Ans = vbYes Then
         End
      Else
         Cancel = 1
         反安裝框計時器.Enabled = True
         Exit Sub
      End If
   End If
   正在安裝檔案計時器.Enabled = False
   Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
      If 第二安裝頁框.Visible = True Then
         正在安裝檔案計時器.Enabled = True
      End If
   End If
End Sub
Private Sub 反安裝框計時器_Timer() ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   CountTimeInstallation = CountTimeInstallation + 1
   If CountTimeInstallation = 1 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\00000404.016"
      反安裝框條型(0).Visible = True
   ElseIf CountTimeInstallation = 10 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\00000404.256"
   ElseIf CountTimeInstallation = 20 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\age2_x1.exe"
   ElseIf CountTimeInstallation = 25 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\AGE2_X1.ICD"
      反安裝框條型(1).Visible = True
      反安裝框條型(2).Visible = True
   ElseIf CountTimeInstallation = 40 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\AOC10.exe"
      反安裝框條型(3).Visible = True
   ElseIf CountTimeInstallation = 50 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\clcd16.dll"
      反安裝框條型(4).Visible = True
   ElseIf CountTimeInstallation = 70 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\clcd32.dll"
   ElseIf CountTimeInstallation = 90 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\dplayerx.dll"
      反安裝框條型(5).Visible = True
   ElseIf CountTimeInstallation = 95 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\drvmgt.dll"
   ElseIf CountTimeInstallation = 100 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EBUEula.dll"
      反安裝框條型(6).Visible = True
   ElseIf CountTimeInstallation = 115 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\ebueulax.dll"
   ElseIf CountTimeInstallation = 120 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EBUSetup.sem"
      反安裝框條型(7).Visible = True
   ElseIf CountTimeInstallation = 135 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\empires2.exe"
   ElseIf CountTimeInstallation = 200 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EULA.RTF"
      反安裝框條型(8).Visible = True
   ElseIf CountTimeInstallation = 210 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EULAx.RTF"
      反安裝框條型(9).Visible = True
   ElseIf CountTimeInstallation = 255 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\HA312W32.DLL"
   ElseIf CountTimeInstallation = 261 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\language.dll"
      反安裝框條型(10).Visible = True
      反安裝框條型(11).Visible = True
   ElseIf CountTimeInstallation = 270 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\language_x1.dll"
   ElseIf CountTimeInstallation = 296 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\language_x1_p1.dll"
      反安裝框條型(12).Visible = True
   ElseIf CountTimeInstallation = 306 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\myth.acm"
   ElseIf CountTimeInstallation = 363 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player9.hki"
      反安裝框條型(13).Visible = True
   ElseIf CountTimeInstallation = 379 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player10.hki"
   ElseIf CountTimeInstallation = 380 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player.nfo"
      反安裝框條型(14).Visible = True
   ElseIf CountTimeInstallation = 382 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player.nfp"
   ElseIf CountTimeInstallation = 384 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player.nfx"
      反安裝框條型(15).Visible = True
   ElseIf CountTimeInstallation = 386 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Readme.rtf"
   ElseIf CountTimeInstallation = 389 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Readmex.rtf"
      反安裝框條型(16).Visible = True
   ElseIf CountTimeInstallation = 399 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\ReadmeX_a.rtf"
   ElseIf CountTimeInstallation = 410 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\scenariobkg.bmp"
      反安裝框條型(17).Visible = True
   ElseIf CountTimeInstallation = 420 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SETUPENU.DLL"
   ElseIf CountTimeInstallation = 445 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SetupReg.exe"
   ElseIf CountTimeInstallation = 450 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SHW32.DLL"
      反安裝框條型(18).Visible = True
   ElseIf CountTimeInstallation = 475 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\STPENUX.DLL"
   ElseIf CountTimeInstallation = 480 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\yi8oytru.dll"
      反安裝框條型(19).Visible = True
   ElseIf CountTimeInstallation = 490 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\rirsegsda.dll"
   ElseIf CountTimeInstallation = 512 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\q35tyrwe.dll"
      反安裝框條型(20).Visible = True
   ElseIf CountTimeInstallation = 515 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\aeghjda.dll"
   ElseIf CountTimeInstallation = 516 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\serhyjykf.dll"
      反安裝框條型(21).Visible = True
   ElseIf CountTimeInstallation = 520 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\sahjsdgh.dll"
   ElseIf CountTimeInstallation = 570 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SaveGame\Save.dll"
      反安裝框條型(22).Visible = True
   ElseIf CountTimeInstallation = 580 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SaveGame\Inf.inf"
      反安裝框條型(23).Visible = True
   ElseIf CountTimeInstallation = 595 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SaveGame\Top.txt"
      反安裝框條型(24).Visible = True
   ElseIf CountTimeInstallation = 600 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Sound\0003.mp3"
      反安裝框條型(25).Visible = True
   ElseIf CountTimeInstallation = 610 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Sound\0006.mp3"
      反安裝框條型(26).Visible = True
   ElseIf CountTimeInstallation = 620 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Sound\Credits.mp3"
      反安裝框條型(27).Visible = True
   ElseIf CountTimeInstallation = 630 Then
      安裝狀態框路徑.Caption = "C:\Program Files\...\Sound\Revolootin.mp3"
      反安裝框條型(28).Visible = True
   ElseIf CountTimeInstallation = 640 Then
      安裝狀態框路徑.Caption = "C:\Program Files\...\Sound\SomeOfAKind.mp3"
      反安裝框條型(29).Visible = True
   ElseIf CountTimeInstallation = 660 Then '完成
      Call Kill_File("C:\WINDOWS\ChineseChess.ini")
      Call Kill_File("C:\Program Files\ChineseChess\Chinese Chess 2.3.0.exe")
      反安裝框計時器.Enabled = False
      Temp = MsgBox("The ChineseChess Has Been Removed.", vbOKOnly + vbInformation, "")
      End
   End If
End Sub
Private Sub 正在安裝檔案計時器_Timer()
   CountTimeInstallation = CountTimeInstallation + 1
   If CountTimeInstallation = 1 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\00000404.016"
      條型(0).Visible = True
   ElseIf CountTimeInstallation = 10 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\00000404.256"
   ElseIf CountTimeInstallation = 20 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\age2_x1.exe"
   ElseIf CountTimeInstallation = 25 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\AGE2_X1.ICD"
      條型(1).Visible = True
      條型(2).Visible = True
   ElseIf CountTimeInstallation = 40 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\AOC10.exe"
      條型(3).Visible = True
   ElseIf CountTimeInstallation = 50 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\clcd16.dll"
      條型(4).Visible = True
   ElseIf CountTimeInstallation = 70 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\clcd32.dll"
   ElseIf CountTimeInstallation = 90 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\dplayerx.dll"
      條型(5).Visible = True
   ElseIf CountTimeInstallation = 95 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\drvmgt.dll"
   ElseIf CountTimeInstallation = 100 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\EBUEula.dll"
      條型(6).Visible = True
   ElseIf CountTimeInstallation = 115 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\ebueulax.dll"
   ElseIf CountTimeInstallation = 120 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\EBUSetup.sem"
      條型(7).Visible = True
   ElseIf CountTimeInstallation = 135 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\empires2.exe"
   ElseIf CountTimeInstallation = 200 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\EULA.RTF"
      條型(8).Visible = True
   ElseIf CountTimeInstallation = 210 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\EULAx.RTF"
      條型(9).Visible = True
   ElseIf CountTimeInstallation = 255 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\HA312W32.DLL"
   ElseIf CountTimeInstallation = 261 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\language.dll"
      條型(10).Visible = True
      條型(11).Visible = True
   ElseIf CountTimeInstallation = 270 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\language_x1.dll"
   ElseIf CountTimeInstallation = 296 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\language_x1_p1.dll"
      條型(12).Visible = True
   ElseIf CountTimeInstallation = 306 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\myth.acm"
   ElseIf CountTimeInstallation = 363 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\player9.hki"
      條型(13).Visible = True
   ElseIf CountTimeInstallation = 379 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\player10.hki"
   ElseIf CountTimeInstallation = 380 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\player.nfo"
      條型(14).Visible = True
   ElseIf CountTimeInstallation = 382 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\player.nfp"
   ElseIf CountTimeInstallation = 384 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\player.nfx"
      條型(15).Visible = True
   ElseIf CountTimeInstallation = 386 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\Readme.rtf"
   ElseIf CountTimeInstallation = 389 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\Readmex.rtf"
      條型(16).Visible = True
   ElseIf CountTimeInstallation = 399 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\ReadmeX_a.rtf"
   ElseIf CountTimeInstallation = 410 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\scenariobkg.bmp"
      條型(17).Visible = True
   ElseIf CountTimeInstallation = 420 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\SETUPENU.DLL"
   ElseIf CountTimeInstallation = 445 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\SetupReg.exe"
   ElseIf CountTimeInstallation = 450 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\SHW32.DLL"
      條型(18).Visible = True
   ElseIf CountTimeInstallation = 475 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\STPENUX.DLL"
   ElseIf CountTimeInstallation = 480 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\yi8oytru.dll"
      條型(19).Visible = True
   ElseIf CountTimeInstallation = 490 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\rirsegsda.dll"
   ElseIf CountTimeInstallation = 512 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\q35tyrwe.dll"
      條型(20).Visible = True
   ElseIf CountTimeInstallation = 515 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\aeghjda.dll"
   ElseIf CountTimeInstallation = 516 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\serhyjykf.dll"
      條型(21).Visible = True
   ElseIf CountTimeInstallation = 520 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\sahjsdgh.dll"
   ElseIf CountTimeInstallation = 570 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\SaveGame\Save.dll"
      條型(22).Visible = True
   ElseIf CountTimeInstallation = 580 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\SaveGame\Inf.inf"
      條型(23).Visible = True
   ElseIf CountTimeInstallation = 595 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\SaveGame\Top.txt"
      條型(24).Visible = True
   ElseIf CountTimeInstallation = 600 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\Sound\0003.mp3"
      條型(25).Visible = True
   ElseIf CountTimeInstallation = 610 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\Sound\0006.mp3"
      條型(26).Visible = True
   ElseIf CountTimeInstallation = 620 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\Chinese Chess\Sound\Credits.mp3"
      條型(27).Visible = True
   ElseIf CountTimeInstallation = 630 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\...\Sound\Revolootin.mp3"
      條型(28).Visible = True
   ElseIf CountTimeInstallation = 640 Then
      第二安裝頁框正在安裝檔案名稱.Caption = "C:\Program Files\...\Sound\SomeOfAKind.mp3"
      條型(29).Visible = True
   ElseIf CountTimeInstallation = 660 Then '完成安裝
   ''
      Call RealInstall '真正安裝
   ''
      Open "C:\WINDOWS\ChineseChess.ini" For Output As #1 '建立檔案
         Print #1, 開始表單.檔案版本數值 '寫入此版本
      Close #1
      For i = 0 To 29
         條型(i).Visible = False
      Next
      正在安裝檔案計時器.Enabled = False
      第二安裝頁框.Visible = False
      第三安裝頁框.Visible = True
   End If
End Sub
Private Sub RealInstall() '真正安裝程式殼
   '寫入確認檔
   'Open "C:\WINDOWS\ChineseChess.ini" For Output As #1 '建立檔案
   '   Print #1, "V2.3" '寫入此版本
   'Close #1
   '執行外部程式
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\FilePath.dll")
   Open "C:\Program Files\MakeFilesPath.txt" For Output As #1 '建立檔案
      Print #1, "ChineseChess" '檔案名稱
   Close #1
   Shell ("C:\Program Files\FilePath.dll"), vbHide
Return0:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '程式運行中
       GoTo Return0
   Else '程式未運行
      Call Kill_File("C:\Program Files\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\FilePath.dll")
   End If
   '
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\ChineseChess\FilePath.dll")
   Open "C:\Program Files\ChineseChess\MakeFilesPath.txt" For Output As #1 '建立檔案
      Print #1, "bin" '檔案名稱
   Close #1
   Shell ("C:\Program Files\ChineseChess\FilePath.dll"), vbHide
Return1:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '程式運行中
       GoTo Return1
   Else '程式未運行
      Call Kill_File("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\ChineseChess\FilePath.dll")
   End If
   '
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\ChineseChess\FilePath.dll")
   Open "C:\Program Files\ChineseChess\MakeFilesPath.txt" For Output As #1 '建立檔案
      Print #1, "Picture" '檔案名稱
   Close #1
   Shell ("C:\Program Files\ChineseChess\FilePath.dll"), vbHide
Return2:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '程式運行中
       GoTo Return2
   Else '程式未運行
      Call Kill_File("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\ChineseChess\FilePath.dll")
   End If
   '
   FileCopy (App.Path & "\OutSideExe\FilePath.dll"), ("C:\Program Files\ChineseChess\FilePath.dll")
   Open "C:\Program Files\ChineseChess\MakeFilesPath.txt" For Output As #1 '建立檔案
      Print #1, "Sound" '檔案名稱
   Close #1
   Shell ("C:\Program Files\ChineseChess\FilePath.dll"), vbHide
Return3:
   StrSQL = "Select * from Win32_Process Where Name = 'FilePath.dll'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then '程式運行中
       GoTo Return2
   Else '程式未運行
      Call Kill_File("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_File("C:\Program Files\ChineseChess\FilePath.dll")
   End If
   '複製檔案
   InstallCount = 0
DoNextInstall:
   InstallCount = InstallCount + 1
   'CanDoNext = False
   Call RealInstallCount
   第二安裝頁框正在安裝檔案名稱.Caption = InstallNewPath & InstallNewFilePath
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
   If InstallCount = 112 Then '安裝完畢
      GoTo EndInstall
   End If
'CheckTimer:
   'If CanDoNext = True Then
      GoTo DoNextInstall
   'Else
   '   GoTo CheckTimer
   'End If
EndInstall:
   '第二安裝頁框.Visible = False
   '第三安裝頁框.Visible = True
End Sub
Private Sub Kill_File(ByRef Dir_String As String)
On Error GoTo Exit_out
Kill Dir_String
Exit_out:
End Sub
Private Sub RealInstallCount() '真正安裝內容
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
      InstallNewPath = "C:\Documents and Settings\All Users\桌面\"
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
Private Sub 安裝完成框完成_Click()
   End
End Sub
Private Sub 安裝狀態框取消_Click()
   安裝狀態框計時器.Enabled = False
   Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      結束框.Visible = True
      安裝狀態框.Visible = False
   Else
      Cancel = 1
      安裝狀態框計時器.Enabled = True
   End If
End Sub
Private Sub 安裝狀態框計時器_Timer()
   CountTimeInstallation = CountTimeInstallation + 1
   If CountTimeInstallation = 1 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\00000404.016"
      安裝狀態框條型(0).Visible = True
   ElseIf CountTimeInstallation = 10 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\00000404.256"
   ElseIf CountTimeInstallation = 20 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\age2_x1.exe"
   ElseIf CountTimeInstallation = 25 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\AGE2_X1.ICD"
      安裝狀態框條型(1).Visible = True
      安裝狀態框條型(2).Visible = True
   ElseIf CountTimeInstallation = 40 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\AOC10.exe"
      安裝狀態框條型(3).Visible = True
   ElseIf CountTimeInstallation = 50 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\clcd16.dll"
      安裝狀態框條型(4).Visible = True
   ElseIf CountTimeInstallation = 70 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\clcd32.dll"
   ElseIf CountTimeInstallation = 90 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\dplayerx.dll"
      安裝狀態框條型(5).Visible = True
   ElseIf CountTimeInstallation = 95 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\drvmgt.dll"
   ElseIf CountTimeInstallation = 100 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EBUEula.dll"
      安裝狀態框條型(6).Visible = True
   ElseIf CountTimeInstallation = 115 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\ebueulax.dll"
   ElseIf CountTimeInstallation = 120 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EBUSetup.sem"
      安裝狀態框條型(7).Visible = True
   ElseIf CountTimeInstallation = 135 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\empires2.exe"
   ElseIf CountTimeInstallation = 200 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EULA.RTF"
      安裝狀態框條型(8).Visible = True
   ElseIf CountTimeInstallation = 210 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\EULAx.RTF"
      安裝狀態框條型(9).Visible = True
   ElseIf CountTimeInstallation = 255 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\HA312W32.DLL"
   ElseIf CountTimeInstallation = 261 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\language.dll"
      安裝狀態框條型(10).Visible = True
      安裝狀態框條型(11).Visible = True
   ElseIf CountTimeInstallation = 270 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\language_x1.dll"
   ElseIf CountTimeInstallation = 296 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\language_x1_p1.dll"
      安裝狀態框條型(12).Visible = True
   ElseIf CountTimeInstallation = 306 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\myth.acm"
   ElseIf CountTimeInstallation = 363 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player9.hki"
      安裝狀態框條型(13).Visible = True
   ElseIf CountTimeInstallation = 379 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player10.hki"
   ElseIf CountTimeInstallation = 380 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player.nfo"
      安裝狀態框條型(14).Visible = True
   ElseIf CountTimeInstallation = 382 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player.nfp"
   ElseIf CountTimeInstallation = 384 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\player.nfx"
      安裝狀態框條型(15).Visible = True
   ElseIf CountTimeInstallation = 386 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Readme.rtf"
   ElseIf CountTimeInstallation = 389 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Readmex.rtf"
      安裝狀態框條型(16).Visible = True
   ElseIf CountTimeInstallation = 399 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\ReadmeX_a.rtf"
   ElseIf CountTimeInstallation = 410 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\scenariobkg.bmp"
      安裝狀態框條型(17).Visible = True
   ElseIf CountTimeInstallation = 420 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SETUPENU.DLL"
   ElseIf CountTimeInstallation = 445 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SetupReg.exe"
   ElseIf CountTimeInstallation = 450 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SHW32.DLL"
      安裝狀態框條型(18).Visible = True
   ElseIf CountTimeInstallation = 475 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\STPENUX.DLL"
   ElseIf CountTimeInstallation = 480 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\yi8oytru.dll"
      安裝狀態框條型(19).Visible = True
   ElseIf CountTimeInstallation = 490 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\rirsegsda.dll"
   ElseIf CountTimeInstallation = 512 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\q35tyrwe.dll"
      安裝狀態框條型(20).Visible = True
   ElseIf CountTimeInstallation = 515 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\aeghjda.dll"
   ElseIf CountTimeInstallation = 516 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\serhyjykf.dll"
      安裝狀態框條型(21).Visible = True
   ElseIf CountTimeInstallation = 520 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\sahjsdgh.dll"
   ElseIf CountTimeInstallation = 570 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SaveGame\Save.dll"
      安裝狀態框條型(22).Visible = True
   ElseIf CountTimeInstallation = 580 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SaveGame\Inf.inf"
      安裝狀態框條型(23).Visible = True
   ElseIf CountTimeInstallation = 595 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\SaveGame\Top.txt"
      安裝狀態框條型(24).Visible = True
   ElseIf CountTimeInstallation = 600 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Sound\0003.mp3"
      安裝狀態框條型(25).Visible = True
   ElseIf CountTimeInstallation = 610 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Sound\0006.mp3"
      安裝狀態框條型(26).Visible = True
   ElseIf CountTimeInstallation = 620 Then
      安裝狀態框路徑.Caption = "C:\Program Files\Chinese Chess\Sound\Credits.mp3"
      安裝狀態框條型(27).Visible = True
   ElseIf CountTimeInstallation = 630 Then
      安裝狀態框路徑.Caption = "C:\Program Files\...\Sound\Revolootin.mp3"
      安裝狀態框條型(28).Visible = True
   ElseIf CountTimeInstallation = 640 Then
      安裝狀態框路徑.Caption = "C:\Program Files\...\Sound\SomeOfAKind.mp3"
      安裝狀態框條型(29).Visible = True
   ElseIf CountTimeInstallation = 660 Then '完成安裝
   ''
      Call RealInstall '真正安裝
   ''
      Open "C:\WINDOWS\ChineseChess.ini" For Output As #1 '建立檔案
         Print #1, 開始表單.檔案版本數值 '寫入此版本
      Close #1
      For i = 0 To 29
         更新框條型(i).Visible = False
      Next
      安裝狀態框計時器.Enabled = False
      更新框.Visible = False
      安裝完成框.Visible = True
   End If
End Sub
Private Sub 更新框取消_Click()
   Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      結束框.Visible = True
      修改修復移除框.Visible = False
   Else
      Cancel = 1
   End If
End Sub
Private Sub 修改修復移除框下一步_Click()
   If 修改修復或移除程式選項(0).Value = True Then '修改
      Call ShowSpaceInfo("C:")
      修改框.Visible = True
      修改修復移除框.Visible = False
   ElseIf 修改修復或移除程式選項(1).Value = True Then '修復
      安裝狀態框.Visible = True
      修改修復移除框.Visible = False
      CountTimeInstallation = 0
      For i = 0 To 29
         If i > 0 Then
            Load 安裝狀態框條型(i)
         End If
         安裝狀態框條型(i).Left = 安裝狀態框條型(0).Left + i * (安裝狀態框條型(0).Width + 10)
         安裝狀態框條型(i).Top = 安裝狀態框條型(0).Top
      Next
      安裝狀態框計時器.Enabled = True
   ElseIf 修改修復或移除程式選項(2).Value = True Then '移除
      反安裝框.Visible = True
      修改修復移除框.Visible = False
      CountTimeInstallation = 0
      For i = 0 To 29
         If i > 0 Then
            Load 反安裝框條型(i)
            '反安裝框條型(i).Visible = True
         End If
         反安裝框條型(i).Left = 反安裝框條型(0).Left + i * (反安裝框條型(0).Width + 10)
         反安裝框條型(i).Top = 反安裝框條型(0).Top
      Next
      反安裝框計時器.Enabled = True
   End If
End Sub
Private Sub 修改修復移除框取消_Click()
   Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      結束框.Visible = True
      修改修復移除框.Visible = False
   Else
      Cancel = 1
   End If
End Sub
Private Sub 修改框MainApp_Click()
   If 修改框MainApp.Value = 0 Then
      修改框磁碟統計需要空間.Caption = "需要 0.00 MB 的空間（在 C 磁碟上）"
   Else
      修改框磁碟統計需要空間.Caption = "需要 29.1 MB 的空間（在 C 磁碟上）"
   End If
End Sub
Private Sub 修改框下一步_Click()
   If 修改框MainApp.Value = 1 Then
      維護完成框.Visible = True
      修改框.Visible = False
   Else
      
   End If
End Sub
Private Sub 修改框上一步_Click()
   修改修復移除框.Visible = True
   修改框.Visible = False
End Sub
Private Sub 修改框取消_Click()
   Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      結束框.Visible = True
      修改框.Visible = False
   Else
      Cancel = 1
   End If
End Sub
Private Sub 第一安裝頁框下一步_Click()
   第一安裝頁框.Visible = False
   第二安裝頁框.Visible = True
   For i = 0 To 29
      If i > 0 Then
         Load 條型(i)
      End If
      條型(i).Left = 條型(0).Left + i * (條型(0).Width + 10)
      條型(i).Top = 條型(0).Top
   Next
   正在安裝檔案計時器.Enabled = True '假安裝
   'Call RealInstall '真正安裝
End Sub
Private Sub 第一安裝頁框取消_Click()
   Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
   End If
End Sub
Private Sub 第二安裝頁框取消_Click()
   正在安裝檔案計時器.Enabled = False
   Ans = MsgBox("確實要取消安裝嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
      正在安裝檔案計時器.Enabled = True
   End If
End Sub
Private Sub 反安裝框取消_Click()
   反安裝框計時器.Enabled = False
   Ans = MsgBox("確實要取消嗎？", vbYesNo + vbExclamation, "結束安裝程式")
   If Ans = vbYes Then
      End
   Else
      Cancel = 1
      反安裝框計時器.Enabled = True
   End If
End Sub
Private Sub 第三安裝頁框完成_Click()
   If 開啟象棋.Value = 1 Then '勾選
      Shell "C:\Program Files\ChineseChess\Chinese Chess 2.3.0.exe", vbNormalFocus
   End If
   End
End Sub
Private Sub 結束框完成_Click()
   End
End Sub
'Private Sub 維護完成框_DragDrop(Source As Control, X As Single, Y As Single)
'維護完成框.ZOrder
'End Sub
Private Sub 維護完成框完成_Click()
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
   修改框磁碟統計剩餘空間.Caption = "可用 " & Sum & " 的空間"
End Sub
