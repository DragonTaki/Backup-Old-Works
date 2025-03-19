VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form 主頁 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋"
   ClientHeight    =   12270
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   17925
   Icon            =   "M04_象棋.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12270
   ScaleWidth      =   17925
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame 關於框 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  '沒有框線
      Height          =   10380
      Left            =   0
      TabIndex        =   28
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.TextBox 日期 
         Appearance      =   0  '平面
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         ToolTipText     =   "關於..."
         Top             =   3360
         Width           =   10695
      End
      Begin VB.TextBox 版本 
         Appearance      =   0  '平面
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         ToolTipText     =   "關於..."
         Top             =   960
         Width           =   10695
      End
      Begin VB.TextBox 連結 
         Appearance      =   0  '平面
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         MousePointer    =   14  '箭號與問號形狀
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "M04_象棋.frx":324A
         ToolTipText     =   "關於..."
         Top             =   4200
         Width           =   10695
      End
      Begin VB.CommandButton 確定 
         BackColor       =   &H006957BD&
         Caption         =   "確　　定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1178
         Style           =   1  '圖片外觀
         TabIndex        =   30
         ToolTipText     =   "確定"
         Top             =   9120
         Width           =   10695
      End
      Begin VB.TextBox 關於文字 
         BackColor       =   &H00E4673D&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
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
         MousePointer    =   1  '箭號形狀
         MultiLine       =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "M04_象棋.frx":327F
         ToolTipText     =   "關於..."
         Top             =   960
         Width           =   10695
      End
   End
   Begin VB.Frame 說明框 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  '沒有框線
      Height          =   10380
      Left            =   0
      TabIndex        =   32
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.CommandButton 說明確定 
         BackColor       =   &H00FFFFC0&
         Caption         =   "確　　定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   20.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         Style           =   1  '圖片外觀
         TabIndex        =   33
         ToolTipText     =   "確定"
         Top             =   8640
         Width           =   5535
      End
      Begin VB.Label 象棋規則 
         Caption         =   $"M04_象棋.frx":334E
         BeginProperty Font 
            Name            =   "新細明體"
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
         ToolTipText     =   "規則"
         Top             =   1800
         Width           =   11775
      End
      Begin VB.Label 說明標題 
         BackColor       =   &H00C0E0FF&
         Caption         =   "　　　遊戲規則"
         BeginProperty Font 
            Name            =   "新細明體"
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
   Begin VB.Frame 主頁框 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  '沒有框線
      Height          =   10380
      Left            =   0
      TabIndex        =   5
      Top             =   -120
      Width           =   13140
      Begin VB.CommandButton 進入遊戲 
         BackColor       =   &H00C0FFFF&
         Caption         =   "進入遊戲"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '圖片外觀
         TabIndex        =   0
         ToolTipText     =   "遊戲開始"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton 遊戲設定 
         BackColor       =   &H00C0FFFF&
         Caption         =   "遊戲設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '圖片外觀
         TabIndex        =   1
         ToolTipText     =   "設定音樂ˋ規則等"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton 關於按鈕 
         BackColor       =   &H00C0FFFF&
         Caption         =   "關　　於"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '圖片外觀
         TabIndex        =   3
         ToolTipText     =   "關於本程式"
         Top             =   6240
         Width           =   1935
      End
      Begin VB.CommandButton 遊戲說明按鈕 
         BackColor       =   &H00C0FFFF&
         Caption         =   "遊戲說明"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '圖片外觀
         TabIndex        =   2
         ToolTipText     =   "說明遊戲規則"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton 結束按鈕 
         BackColor       =   &H00C0FFFF&
         Caption         =   "結　　束"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   18
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8880
         Style           =   1  '圖片外觀
         TabIndex        =   4
         ToolTipText     =   "結束遊戲"
         Top             =   8160
         Width           =   1935
      End
      Begin VB.Image 主頁圖片 
         Height          =   9330
         Left            =   1080
         Picture         =   "M04_象棋.frx":35D3
         Stretch         =   -1  'True
         ToolTipText     =   "中國象棋"
         Top             =   480
         Width           =   6360
      End
   End
   Begin VB.Frame 設定框 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      Height          =   10395
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.CommandButton 回主頁 
         BackColor       =   &H00FFFFFF&
         Caption         =   "回主頁"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   8160
         Width           =   6135
      End
      Begin VB.Frame 炮跳 
         BackColor       =   &H00000000&
         Caption         =   "炮跳"
         BeginProperty Font 
            Name            =   "新細明體"
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
         ToolTipText     =   "炮是否用跳吃"
         Top             =   2640
         Width           =   6135
         Begin VB.OptionButton 炮跳關 
            BackColor       =   &H00000000&
            Caption         =   "關"
            BeginProperty Font 
               Name            =   "新細明體"
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
            ToolTipText     =   "設定炮跳關"
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton 炮跳開 
            BackColor       =   &H00000000&
            Caption         =   "開"
            BeginProperty Font 
               Name            =   "新細明體"
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
            ToolTipText     =   "設定炮跳開"
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame 音樂 
         BackColor       =   &H00000000&
         Caption         =   "音樂"
         BeginProperty Font 
            Name            =   "新細明體"
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
         ToolTipText     =   "設定音樂開關"
         Top             =   5400
         Width           =   6135
         Begin VB.Timer 音樂計時器 
            Interval        =   1000
            Left            =   3480
            Top             =   600
         End
         Begin VB.OptionButton 音樂關 
            BackColor       =   &H00000000&
            Caption         =   "關"
            BeginProperty Font 
               Name            =   "新細明體"
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
            ToolTipText     =   "設定音樂關"
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton 音樂開 
            BackColor       =   &H00000000&
            Caption         =   "開"
            BeginProperty Font 
               Name            =   "新細明體"
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
            ToolTipText     =   "設定音樂開"
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin MCI.MMControl 撥放器 
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
      Begin VB.Image 設定圖 
         Height          =   6015
         Left            =   360
         Picture         =   "M04_象棋.frx":E83E
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   5985
      End
      Begin VB.Label 設定 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00000000&
         Caption         =   "設定"
         BeginProperty Font 
            Name            =   "新細明體"
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
   Begin VB.Frame 重新框 
      BackColor       =   &H005FA76F&
      BorderStyle     =   0  '沒有框線
      Height          =   10380
      Left            =   0
      TabIndex        =   36
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.Timer 結束圖片計時器 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   0
         Top             =   0
      End
      Begin VB.Image 結束圖片 
         Height          =   1215
         Left            =   3960
         Picture         =   "M04_象棋.frx":12E17
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label 結束訊息 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 按任意鍵重新 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  '透明
         Caption         =   "按任意鍵重新"
         BeginProperty Font 
            Name            =   "新細明體"
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
   Begin VB.Frame 象棋框 
      BackColor       =   &H005FA76F&
      BorderStyle     =   0  '沒有框線
      Caption         =   "s"
      Height          =   10380
      Left            =   0
      TabIndex        =   16
      Top             =   -120
      Visible         =   0   'False
      Width           =   13140
      Begin VB.Frame 象棋圖框 
         BorderStyle     =   0  '沒有框線
         Height          =   1335
         Left            =   11760
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image 空白圖 
            Height          =   1215
            Left            =   0
            Picture         =   "M04_象棋.frx":12F7A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   31
            Left            =   0
            Picture         =   "M04_象棋.frx":130DD
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   30
            Left            =   0
            Picture         =   "M04_象棋.frx":1361D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   29
            Left            =   0
            Picture         =   "M04_象棋.frx":13DEC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   28
            Left            =   0
            Picture         =   "M04_象棋.frx":1432C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   27
            Left            =   0
            Picture         =   "M04_象棋.frx":1486C
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   26
            Left            =   0
            Picture         =   "M04_象棋.frx":14DAC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   25
            Left            =   0
            Picture         =   "M04_象棋.frx":15336
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   24
            Left            =   0
            Picture         =   "M04_象棋.frx":158C0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   23
            Left            =   0
            Picture         =   "M04_象棋.frx":15E36
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   22
            Left            =   0
            Picture         =   "M04_象棋.frx":163AC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   21
            Left            =   0
            Picture         =   "M04_象棋.frx":1695B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   20
            Left            =   0
            Picture         =   "M04_象棋.frx":16F0A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   19
            Left            =   0
            Picture         =   "M04_象棋.frx":173E6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   18
            Left            =   0
            Picture         =   "M04_象棋.frx":178C2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   17
            Left            =   0
            Picture         =   "M04_象棋.frx":17DDB
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   16
            Left            =   0
            Picture         =   "M04_象棋.frx":182F4
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   15
            Left            =   0
            Picture         =   "M04_象棋.frx":18823
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   14
            Left            =   0
            Picture         =   "M04_象棋.frx":18D1F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   13
            Left            =   0
            Picture         =   "M04_象棋.frx":1921B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   12
            Left            =   0
            Picture         =   "M04_象棋.frx":19717
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   11
            Left            =   0
            Picture         =   "M04_象棋.frx":19C13
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   10
            Left            =   0
            Picture         =   "M04_象棋.frx":1A10F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   9
            Left            =   0
            Picture         =   "M04_象棋.frx":1A666
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   8
            Left            =   0
            Picture         =   "M04_象棋.frx":1ABBD
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   7
            Left            =   0
            Picture         =   "M04_象棋.frx":1B10E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   6
            Left            =   0
            Picture         =   "M04_象棋.frx":1B65F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   5
            Left            =   0
            Picture         =   "M04_象棋.frx":1BAF9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   4
            Left            =   0
            Picture         =   "M04_象棋.frx":1BF93
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   3
            Left            =   0
            Picture         =   "M04_象棋.frx":1C4DB
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   2
            Left            =   0
            Picture         =   "M04_象棋.frx":1CA23
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   1
            Left            =   0
            Picture         =   "M04_象棋.frx":1CE66
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
         Begin VB.Image 象棋圖 
            Height          =   1215
            Index           =   0
            Left            =   0
            Picture         =   "M04_象棋.frx":1D2A9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.Frame 持棋 
         BackColor       =   &H005FA76F&
         Caption         =   "持棋"
         BeginProperty Font 
            Name            =   "新細明體"
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
         ToolTipText     =   "正要移動的棋"
         Top             =   7440
         Width           =   3255
         Begin VB.Image 持棋圖片 
            Height          =   1215
            Left            =   1200
            Picture         =   "M04_象棋.frx":1D807
            Stretch         =   -1  'True
            Tag             =   "-1"
            ToolTipText     =   "正要移動的棋"
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Frame 剩餘棋數 
         BackColor       =   &H005FA76F&
         Caption         =   "剩餘棋數"
         BeginProperty Font 
            Name            =   "新細明體"
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
         ToolTipText     =   "雙方剩下的棋"
         Top             =   7440
         Width           =   5055
         Begin VB.Label 紅方剩棋 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
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
            ToolTipText     =   "紅棋剩棋數"
            Top             =   720
            Width           =   975
         End
         Begin VB.Label 黑方剩棋 
            Alignment       =   2  '置中對齊
            BackStyle       =   0  '透明
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
            ToolTipText     =   "黑棋剩棋數"
            Top             =   720
            Width           =   975
         End
         Begin VB.Image 帥 
            Height          =   1215
            Left            =   2520
            Picture         =   "M04_象棋.frx":1D96A
            Stretch         =   -1  'True
            Tag             =   "-1"
            ToolTipText     =   "紅棋"
            Top             =   480
            Width           =   1080
         End
         Begin VB.Image 將 
            Height          =   1215
            Left            =   240
            Picture         =   "M04_象棋.frx":1DE99
            Stretch         =   -1  'True
            Tag             =   "-1"
            ToolTipText     =   "黑棋"
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Label 返回 
         Alignment       =   2  '置中對齊
         BackColor       =   &H005FA76F&
         Caption         =   "返回"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 棄權 
         Alignment       =   2  '置中對齊
         BackColor       =   &H005FA76F&
         Caption         =   "棄權"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 重新開始 
         Alignment       =   2  '置中對齊
         BackColor       =   &H005FA76F&
         Caption         =   "重新開始"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Image 換玩家二圖 
         Height          =   570
         Left            =   7125
         Picture         =   "M04_象棋.frx":1E3F7
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Image 換玩家一圖 
         Height          =   570
         Left            =   1125
         Picture         =   "M04_象棋.frx":1F352
         Stretch         =   -1  'True
         Top             =   720
         Width           =   555
      End
      Begin VB.Label 玩家一姓名 
         Alignment       =   2  '置中對齊
         BackColor       =   &H005FA76F&
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 玩家二 
         Alignment       =   2  '置中對齊
         BackColor       =   &H005FA76F&
         Caption         =   "玩家二"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 玩家一 
         Alignment       =   2  '置中對齊
         BackColor       =   &H005FA76F&
         Caption         =   "玩家一"
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   8
         X1              =   11805
         X2              =   11805
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   7
         X1              =   10485
         X2              =   10485
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   6
         X1              =   9165
         X2              =   9165
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   5
         X1              =   7845
         X2              =   7845
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         X1              =   6525
         X2              =   6525
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   3
         X1              =   5205
         X2              =   5205
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   3885
         X2              =   3885
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   2565
         X2              =   2565
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 直線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   1200
         X2              =   1200
         Y1              =   1800
         Y2              =   7080
      End
      Begin VB.Line 橫線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         X1              =   1245
         X2              =   11805
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line 橫線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   3
         X1              =   1245
         X2              =   11805
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line 橫線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   1245
         X2              =   11805
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line 橫線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   1245
         X2              =   11805
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line 橫線 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   1245
         X2              =   11805
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Image 棋 
         Height          =   1215
         Index           =   0
         Left            =   1365
         Picture         =   "M04_象棋.frx":202AD
         Stretch         =   -1  'True
         Tag             =   "-1"
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label 玩家二姓名 
         Alignment       =   2  '置中對齊
         BackColor       =   &H005FA76F&
         BeginProperty Font 
            Name            =   "新細明體"
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
   Begin VB.Menu 檔案 
      Caption         =   "檔案(&F)"
      Begin VB.Menu 結束 
         Caption         =   "結束(&X)"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu 說明 
      Caption         =   "說明(&H)"
      Begin VB.Menu 遊戲說明 
         Caption         =   "遊戲說明(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu 關於 
         Caption         =   "關於(&A)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "主頁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Name1P As String
Dim Name2P As String
Dim 是否炮跳 As Boolean
Dim 是否音樂 As Boolean
Dim 須安裝 As Boolean
Dim CountTime As Integer
Dim 已按第一顆棋 As Boolean
Dim 已開局 As Boolean '選顏色用
Dim 棋已翻開(32) As Boolean '用位子記
Dim 空格(32) As Boolean '用位子記
Dim 現在換玩家二 As Boolean
Dim 玩家一是黑的 As Boolean
Dim 第一顆是自己的棋 As Boolean
Dim 第二顆是自己的棋 As Boolean
Dim 是可移動的 As Boolean
Dim 玩家一勝利 As Boolean
Dim 黑棋勝利 As Boolean
Dim 由右上角結束 As Boolean
Dim 已按返回 As Boolean
Dim 第一顆棋 As Integer
Dim 第二顆棋 As Integer
Dim 第一顆棋位置 As Integer
Dim 第二顆棋位置 As Integer
Dim 中間有棋 As Integer
Dim 結束圖片移動距離X As Integer '移動方向變數
Dim 結束圖片移動距離Y As Integer
Dim 結束訊息移動距離X As Integer
Dim 結束訊息移動距離Y As Integer
Dim Temp1 As Integer
Dim Temp2 As String
Dim 音樂計時器CountTime As Integer
Private Sub Form_Load()
   版本.Text = "象棋 " & 開始表單.檔案版本
   日期.Text = "日期：2012.12.11."
   主頁.Height = 10780
   主頁.Width = 13140
   主頁框.Left = 0
   主頁框.Top = 0
   主頁框.Height = 10380
   主頁框.Width = 18015
   象棋框.Left = 0
   象棋框.Top = 0
   象棋框.Height = 10380
   象棋框.Width = 18015
   重新框.Left = 0
   重新框.Top = 0
   重新框.Height = 10380
   重新框.Width = 18015
   說明框.Left = 0
   說明框.Top = 0
   說明框.Height = 10380
   說明框.Width = 18015
   關於框.Left = 0
   關於框.Top = 0
   關於框.Height = 10380
   關於框.Width = 18015
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim TheFile As String
   Dim Results As String
   TheFile = "C:\WINDOWS\ChineseChess.ini"
   Results = Dir$(TheFile)
   If Results = "" Then '檔案不存在
      FileCopy (App.Path & "\象棋" & 開始表單.檔案版本 & "安裝檔.exe"), ("C:\WINDOWS\象棋" & 開始表單.檔案版本 & ".exe")
      FileCopy (App.Path & "\空白.che"), ("C:\WINDOWS\空白.che")
      FileCopy (App.Path & "\0003.mp3"), ("C:\WINDOWS\0003.mp3")
      FileCopy (App.Path & "\0006.mp3"), ("C:\WINDOWS\0006.mp3")
      FileCopy (App.Path & "\Credits.mp3"), ("C:\WINDOWS\Credits.mp3")
      FileCopy (App.Path & "\Revolootin.mp3"), ("C:\WINDOWS\Revolootin.mp3")
      FileCopy (App.Path & "\SomeOfAKind.mp3"), ("C:\WINDOWS\SomeOfAKind.mp3")
      For i = 0 To 31
         FileCopy (App.Path & "\" & i & ".che"), ("C:\WINDOWS\" & i & ".che")
      Next
      須安裝 = True
      開始計時器.Enabled = False
      安裝頁.Show
      Unload 開始表單
   End If
   'SetAttr "C:\WINDOWS\ChineseChess.ini", vbHidden '隱藏
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   玩家一姓名.Caption = Name1P
   玩家二姓名.Caption = Name2P
   Randomize
   '生棋
   For i = 0 To 31
      If i > 0 Then
         Load 棋(i)
      End If
      棋(i).Left = 棋(0).Left + (i Mod 8) * (棋(0).Width + 240) '以0號位置為準，取8的餘數(代表要右偏幾格的寬度+間縫大小)
      棋(i).Top = 棋(0).Top + Int(i / 8) * (棋(0).Height + 105) '以0號位置為準，取整除8(代表要往下偏幾格的高度+間縫大小)
      棋(i).Visible = True
   Next
   Call 洗牌
   結束圖片移動距離X = 150 '把預定移動的速度在此決定，其在X代表左右，若為正，代表開始往右，若為負，代表起動時往左，其值越大，代表移動越快
   結束圖片移動距離Y = -100 '其在Y代表上下，若為正，代表開始往下，若為負，代表起動時往上，其值越大，代表移動越快
   結束訊息移動距離X = 150
   結束訊息移動距離Y = 100
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   設定框.Left = 0
   設定框.Top = 0
   Randomize
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   是否炮跳 = True
   是否音樂 = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MsgBox ("遊戲結束！")
   End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''主頁
Private Sub 結束_Click()
   MsgBox ("遊戲結束！")
   End
End Sub

Private Sub 遊戲說明_Click()
   說明框.Visible = True
   主頁框.Visible = False
End Sub
Private Sub 關於_Click()
   關於框.Visible = True
   主頁框.Visible = False
   Call 隱藏按鈕
End Sub
Private Sub 進入遊戲_Click()
   Name1P = InputBox("輸入第一位玩家姓名", "", "無名")
   If Name1P <> "" Then
      Name2P = InputBox("輸入第二位玩家姓名", "", "無名")
      If Name2P <> "" Then
         象棋框.Visible = True
         主頁框.Visible = False
         Call 隱藏按鈕
      End If
   End If
End Sub
Private Sub 遊戲設定_Click()
   Call 隱藏按鈕
   設定框.Visible = True
   主頁框.Visible = False
End Sub
Private Sub 遊戲說明按鈕_Click()
   Call 隱藏按鈕
   Call 遊戲說明_Click
End Sub
Private Sub 關於按鈕_Click()
   Call 隱藏按鈕
   Call 關於_Click
End Sub
Private Sub 結束按鈕_Click()
    Call 結束_Click
End Sub
Private Sub 隱藏按鈕()
   說明.Visible = False
   檔案.Visible = False
End Sub
Private Sub 顯示按鈕()
   說明.Visible = True
   檔案.Visible = True
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''象棋
Private Sub 返回_Click()
   Ans = MsgBox("你確定要結束？", vbYesNo, "")
   If Ans = vbYes Then
      Name1P = Empty
      Name2P = Empty
      已按第一顆棋 = False
      已開局 = False
      For i = 0 To 32
         棋已翻開(i) = False
      Next
      For j = 0 To 32
         空格(j) = False
      Next
      現在換玩家二 = False
      玩家一是黑的 = False
      第一顆是自己的棋 = False
      第二顆是自己的棋 = False
      是可移動的 = False
      玩家一勝利 = False
      黑棋勝利 = False
      第一顆棋 = Empty
      第二顆棋 = Empty
      第一顆棋位置 = Empty
      第二顆棋位置 = Empty
      中間有棋 = Empty
      由右上角結束 = False
      已按返回 = True
      主頁框.Visible = True
      象棋框.Visible = False
      Call 顯示按鈕
   Else
      Cancel = 1
   End If
End Sub
Private Sub 重新框_Click()
   結束圖片計時器.Enabled = False
   象棋框.Visible = True
   重新框.Visible = False
   Call 重新開始_Click
End Sub
Private Sub 結束圖片計時器_Timer()
   結束圖片.Left = 結束圖片.Left + 結束圖片移動距離X '移動X座標，物件的X標在<Left>這屬性內，其值越大代表越右方
   結束圖片.Top = 結束圖片.Top + 結束圖片移動距離Y '移動Y座標，物件的X標在<Top>這屬性內，其值越大，代表越下方
   If 結束圖片.Left + 結束圖片.Width > 主頁.Width Then '檢查是否超過表單右界(Form1.Width就是表單寬度)，但一個物件的左界值是<Left>，右界值是<Left>+<Width>
      結束圖片移動距離X = 0 - Abs(結束圖片移動距離X) '如果超過移動方向改成負值，代表須向左走
   End If
   If 結束圖片.Top + 結束圖片.Height > 主頁.Height - 600 Then '檢查是否超過表單下界(Form1.Height就是表單高度，代是含藍色標題區，所以要減600，使得到真的下界值)，
      結束圖片移動距離Y = 0 - Abs(結束圖片移動距離Y) '但一個物件的上界值是<Top>，下界值是<top>+<Height>,超過就把移動值改成負的，代表往上走
   End If 'Abs()取一個數的對值
   If 結束圖片.Left < 0 Then '物件一律相對表單移動，所以如果左界值<Left>小於0，代表撞到左牆了，就把值換成正的，改往右走
      結束圖片移動距離X = Abs(結束圖片移動距離X)
   End If
   If 結束圖片.Top < 0 Then '物件一律相對表單移動，所以如果上界值<Top>小於0，代表撞到上牆了，就把值換成正的，改往下走
      結束圖片移動距離Y = Abs(結束圖片移動距離Y)
   End If
   ''''''''''''''''''''''''''''''''''''''''
   結束訊息.Left = 結束訊息.Left + 結束訊息移動距離X '移動X座標，物件的X標在<Left>這屬性內，其值越大代表越右方
   結束訊息.Top = 結束訊息.Top + 結束訊息移動距離Y '移動Y座標，物件的X標在<Top>這屬性內，其值越大，代表越下方
   If 結束訊息.Left + 結束訊息.Width > 主頁.Width Then '檢查是否超過表單右界(Form1.Width就是表單寬度)，但一個物件的左界值是<Left>，右界值是<Left>+<Width>
      結束訊息移動距離X = 0 - Abs(結束訊息移動距離X) '如果超過移動方向改成負值，代表須向左走
   End If
   If 結束訊息.Top + 結束訊息.Height > 主頁.Height - 600 Then '檢查是否超過表單下界(Form1.Height就是表單高度，代是含藍色標題區，所以要減600，使得到真的下界值)，
      結束訊息移動距離Y = 0 - Abs(結束訊息移動距離Y) '但一個物件的上界值是<Top>，下界值是<top>+<Height>,超過就把移動值改成負的，代表往上走
   End If  'Abs()取一個數的對值
   If 結束訊息.Left < 0 Then '物件一律相對表單移動，所以如果左界值<Left>小於0，代表撞到左牆了，就把值換成正的，改往右走
      結束訊息移動距離X = Abs(結束訊息移動距離X)
   End If
   If 結束訊息.Top < 0 Then '物件一律相對表單移動，所以如果上界值<Top>小於0，代表撞到上牆了，就把值換成正的，改往下走
      結束訊息移動距離Y = Abs(結束訊息移動距離Y)
   End If
End Sub
Private Sub 洗牌()
   Dim 棋已發過(32) As Boolean
   For i = 0 To 31
      棋已發過(i) = False
   Next
   For i = 0 To 31
重選棋:
      要放的棋 = Int(Rnd * 32)
      If 棋已發過(要放的棋) = True Then
         GoTo 重選棋
      End If
      棋已發過(要放的棋) = True
      棋(i).Tag = 要放的棋
   Next
End Sub
Private Sub 棋_Click(Index As Integer)
   If 已開局 = False Then '決定顏色
      If 16 > 棋(Index).Tag And 棋(Index).Tag >= 0 Then '點到黑棋
         玩家一.ForeColor = vbBlack
         玩家二.ForeColor = vbRed
         玩家一是黑的 = True
      Else '點到紅棋
         玩家一.ForeColor = vbRed
         玩家二.ForeColor = vbBlack
      End If
      Call 換人
      已開局 = True
      棋(Index).Picture = 象棋圖(棋(Index).Tag).Picture
      棋已翻開(Index) = True
   Else '已決定顏色
      If 棋已翻開(Index) = False Then '該棋未翻開
         棋(Index).Picture = 象棋圖(棋(Index).Tag).Picture
         棋已翻開(Index) = True
         Call 換人
      ElseIf 已按第一顆棋 = True And 棋已翻開(Index) = False Then '第二顆棋未翻開
         If 是否暗吃 = True Then '可暗吃
            GoTo 可暗吃 '視同已翻開
         Else '不可暗吃，離開
            Exit Sub
         End If
         '還沒按
      ElseIf 已按第一顆棋 = False Then '該棋已翻開，是第一顆
         已按第一顆棋 = True
         第一顆棋 = 棋(Index).Tag
         第一顆棋位置 = Index
         Call 判斷第一顆是否為自己的棋
         If 第一顆是自己的棋 = False Then '不是自己的棋
            已按第一顆棋 = False '還沒按
            第一顆棋 = Empty
            第一顆棋位置 = Empty
         Else '是自己的棋
            第一顆是自己的棋 = False
            持棋圖片.Picture = 象棋圖(棋(Index).Tag).Picture
         End If
      ElseIf 已按第一顆棋 = True And 棋已翻開(Index) = True Then '該棋已翻開，是第二顆
         If 棋(Index).Tag = 第一顆棋 Then '和第一顆棋相同，取消持棋
            已按第一顆棋 = False '還沒按
            第一顆棋 = Empty
            第一顆棋位置 = Empty
            持棋圖片.Picture = 空白圖.Picture
         End If
         'If 第一顆棋 = 9 Or 第一顆棋 = 10 Then '例外：第一顆棋是包
         '   If 32 > 棋(Index).Tag And 棋(Index).Tag >= 27 Then '第二顆棋是兵
         '       Exit Sub '視同未按
         '   End If
         'ElseIf 第一顆棋 = 25 Or 第一顆棋 = 26 Then '例外：第一顆棋是炮
         '   If 16 > 棋(Index).Tag And 棋(Index).Tag >= 15 Then '第二顆棋是卒
         '       Exit Sub '視同未按
         '   End If
         'End If
可暗吃:
         第二顆棋 = 棋(Index).Tag
         第二顆棋位置 = Index
         Call 判斷第二顆是否為自己的棋
         If 第二顆是自己的棋 = True Then
            第二顆是自己的棋 = False
            第二顆棋 = Empty
            第二顆棋位置 = Empty
            Exit Sub '第二顆是自己的棋，不能吃，離開
         Else '第二顆不是自己的棋，可以吃
            Call 檢查兩顆棋的位置
            If 是可移動的 = False Then
               Exit Sub
            End If
            是可移動的 = False
            中間有棋 = 0
            Call 動棋子
         End If
      End If
   End If
   If Val(紅方剩棋.Caption) = 0 Then '結束遊戲，黑棋勝利
      If 玩家一是黑的 = True Then '玩家一勝利
         玩家一勝利 = True
         黑棋勝利 = True
         Call 結束遊戲
      Else '玩家二勝利
         玩家一勝利 = False
         黑棋勝利 = True
         Call 結束遊戲
      End If
   ElseIf Val(黑方剩棋.Caption) = 0 Then '結束遊戲，紅棋勝利
      If 玩家一是黑的 = False Then '玩家一勝利
         玩家一勝利 = True
         黑棋勝利 = False
         Call 結束遊戲
      Else '玩家二勝利
         玩家一勝利 = False
         黑棋勝利 = False
         Call 結束遊戲
      End If
   End If
End Sub
Private Sub 判斷第一顆是否為自己的棋()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 16 > 第一顆棋 And 第一顆棋 >= 0 Then '第一顆是黑棋
            第一顆是自己的棋 = True
         End If
      Else '玩家一是紅的
         If 32 > 第一顆棋 And 第一顆棋 >= 16 Then '第一顆是紅棋
            第一顆是自己的棋 = True
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 32 > 第一顆棋 And 第一顆棋 >= 16 Then '第一顆是紅棋
            第一顆是自己的棋 = True
         End If
      Else '玩家二是黑的
         If 16 > 第一顆棋 And 第一顆棋 >= 0 Then '第一顆是黑棋
            第一顆是自己的棋 = True
         End If
      End If
   End If
End Sub
Private Sub 判斷第二顆是否為自己的棋()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '第二顆是黑棋
            第二顆是自己的棋 = True
         End If
      Else '玩家一是紅的
         If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '第二顆是紅棋
            第二顆是自己的棋 = True
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '第二顆是紅棋
            第二顆是自己的棋 = True
         End If
      Else '玩家二是黑的
         If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '第二顆是黑棋
            第二顆是自己的棋 = True
         End If
      End If
   End If
End Sub
Private Sub 換人()
      If 現在換玩家二 = False Then
         現在換玩家二 = True
         換玩家一圖.Visible = False
         換玩家二圖.Visible = True
      Else
         現在換玩家二 = False
         換玩家一圖.Visible = True
         換玩家二圖.Visible = False
      End If
End Sub
Private Sub 檢查兩顆棋的位置()
   If 是否炮跳 = False Then '炮不跳
      GoTo 不開炮跳 '和一般棋動法一樣
   End If
   If 第一顆棋 = 9 Or 第一顆棋 = 10 Or 第一顆棋 = 25 Or 第一顆棋 = 26 Then  '第一顆棋是包或炮，在此移動
      If Abs((第一顆棋位置 - 第二顆棋位置)) < 8 Then '第一顆棋和第二顆棋同行
         If 棋(第二顆棋位置).Tag = -1 Then '第二顆棋是空的
            If 第一顆棋位置 = 第二顆棋位置 + 1 Or 第一顆棋位置 = 第二顆棋位置 - 1 Then '第二顆棋是空的且在第一顆棋上方或下方
               Call 動炮跳清除
            End If
         Else '第二顆棋不是空的
            Temp1 = 第一顆棋位置
            Temp2 = 第二顆棋位置
            If Temp1 > Temp2 Then
               Temp1 = 第二顆棋位置
               Temp2 = 第一顆棋位置
            End If
            For i = Temp1 + 1 To Temp2
               If 棋(i).Tag <> -1 Then
                  中間有棋 = 中間有棋 + 1
               End If
            Next
            If 中間有棋 = 2 Then
               Call 動炮跳清除
            End If
         End If
      ElseIf Abs(第一顆棋位置 - 第二顆棋位置) Mod 8 = 0 Then  '第一顆棋和第二顆棋同列
         If 棋(第二顆棋位置).Tag = -1 Then '第二顆棋是空的
            If 第一顆棋位置 = 第二顆棋位置 + 8 Or 第一顆棋位置 = 第二顆棋位置 - 8 Then '
               Call 動炮跳清除
            End If
         Else '第二顆棋不是空的
            Temp1 = 第一顆棋位置
            Temp2 = 第二顆棋位置
            If Temp1 > Temp2 Then
               Temp1 = 第二顆棋位置
               Temp2 = 第一顆棋位置
            End If
            For i = Temp1 + 8 To Temp2 Step 8
               If 棋(i).Tag <> -1 Then
                  中間有棋 = 中間有棋 + 1
               End If
            Next
            If 中間有棋 = 2 Then
               Call 動炮跳清除
            End If
         End If
      End If
   Else '第一顆棋不是包或炮
不開炮跳:
      If 第一顆棋位置 Mod 8 = 第二顆棋位置 Mod 8 And 第一顆棋位置 = 第二顆棋位置 + 8 Then '第二顆棋在第一顆棋上方
         是可移動的 = True
      ElseIf 第一顆棋位置 Mod 8 = 第二顆棋位置 Mod 8 And 第一顆棋位置 = 第二顆棋位置 - 8 Then '第二顆棋在第一顆棋下方
         是可移動的 = True
      ElseIf 第一顆棋位置 = 第二顆棋位置 + 1 Then '第二顆棋在第一顆棋右邊
         是可移動的 = True
      ElseIf 第一顆棋位置 = 第二顆棋位置 - 1 Then '第二顆棋在第一顆棋左邊
         是可移動的 = True
      End If
   End If
End Sub
Private Sub 吃完清除()
   已按第一顆棋 = False
   第一顆棋 = Empty
   第二顆棋 = Empty
   第一顆棋位置 = Empty
   第二顆棋位置 = Empty
   持棋圖片.Picture = 空白圖.Picture
   Call 換人
End Sub
Private Sub 動炮跳清除()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         End If
      Else '玩家一是紅的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         End If
      Else '玩家二是黑的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         End If
      End If
   End If
   棋(第二顆棋位置).Picture = 象棋圖(第一顆棋).Picture
   棋(第一顆棋位置).Picture = LoadPicture("")
   持棋圖片.Picture = 空白圖.Picture
   Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
   空格(第一顆棋位置) = True
   第一顆棋 = Empty
   第二顆棋 = Empty
   第一顆棋位置 = Empty
   第二顆棋位置 = Empty
   已按第一顆棋 = False
   中間有棋 = 0
   Call 換人
End Sub
Private Sub 動棋子()
   If 現在換玩家二 = False Then
      If 玩家一是黑的 = True Then
         Call 動黑棋
      Else '玩家一是紅的
         Call 動紅棋
      End If
   Else    '現在換玩家二
      If 玩家一是黑的 = True Then
         Call 動紅棋
      Else '玩家二是黑的
         Call 動黑棋
      End If
   End If
End Sub
Private Sub 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag()
   棋(第二顆棋位置).Tag = 棋(第一顆棋位置).Tag
   棋(第一顆棋位置).Tag = -1
End Sub
Private Sub 移棋子一至棋子二()
   棋(第二顆棋位置).Picture = 象棋圖(第一顆棋).Picture
   棋(第一顆棋位置).Picture = LoadPicture("")
End Sub
Private Sub 動黑棋()
   If 第一顆棋 = 0 Then '將
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 27 > 第二顆棋 And 第二顆棋 >= 16 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 27 Then '將吃兵
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 1 Or 第一顆棋 = 2 Then '士
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 17 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 16 Then '士吃帥
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 3 Or 第一顆棋 = 4 Then '象
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 19 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Then '象吃帥仕
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 5 Or 第一顆棋 = 6 Then '車
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 21 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 21 > 第二顆棋 And 第二顆棋 >= 16 Then '車吃帥仕相
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 7 Or 第一顆棋 = 8 Then '馬
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 23 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 23 > 第二顆棋 And 第二顆棋 >= 16 Then '馬吃帥仕相硨
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 9 Or 第一顆棋 = 10 Then '包
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 27 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 27 > 第二顆棋 And 第二顆棋 >= 16 Then
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 16 > 第一顆棋 And 第一顆棋 >= 11 Then '卒
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 25 Or 第二顆棋 = 26 Then '吃炮
         '不動
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 25 > 第二顆棋 And 第二顆棋 >= 17 Then    '卒吃仕相硨碼
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   End If
End Sub
Private Sub 動紅棋()
   If 第一顆棋 = 16 Then '帥
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 11 > 第二顆棋 And 第二顆棋 >= 0 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 15 > 第二顆棋 And 第二顆棋 > 11 Then '帥吃卒
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 17 Or 第一顆棋 = 18 Then '仕
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then   '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Then '仕吃將
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 19 Or 第一顆棋 = 20 Then '相
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Then '相吃將士
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 21 Or 第一顆棋 = 22 Then '硨
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Then '相吃將士象
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 23 Or 第一顆棋 = 24 Then '碼
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Then '相吃將士象車
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 25 Or 第一顆棋 = 26 Then '炮
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Then
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 32 > 第一顆棋 And 第一顆棋 >= 27 Then '兵
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 9 Or 第二顆棋 = 10 Then '吃包
         '不動
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".che")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Then '兵吃士象車馬
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   End If
End Sub
Private Sub 結束遊戲()
   結束圖片計時器.Enabled = True
   重新框.Visible = True
   象棋框.Visible = False
   If 玩家一勝利 = True Then
      結束訊息.Caption = 玩家一姓名 & "勝利"
   Else
      結束訊息.Caption = 玩家二姓名 & "勝利"
   End If
End Sub
Private Sub 重新開始_Click()
   Call 洗牌
   For i = 0 To 31 '蓋牌
      棋(i).Picture = 空白圖.Picture
   Next
   黑方剩棋.Caption = "16"
   紅方剩棋.Caption = "16"
   換玩家一圖.Visible = True
   換玩家二圖.Visible = False
   現在換玩家二 = False
   已按第一顆棋 = False
   For i = 0 To 32
      棋已翻開(i) = False
   Next
   For j = 0 To 32
      空格(j) = False
   Next
   已開局 = False
   現在換玩家二 = False
   玩家一是黑的 = False
   第一顆是自己的棋 = False
   第二顆是自己的棋 = False
   是可移動的 = False
   玩家一勝利 = False
   黑棋勝利 = False
   第一顆棋 = Empty
   第二顆棋 = Empty
   第一顆棋位置 = Empty
   第二顆棋位置 = Empty
   中間有棋 = Empty
   持棋圖片.Picture = 空白圖.Picture
End Sub
Private Sub 棄權_Click()
   If 現在換玩家二 = False Then '玩家一棄權
      玩家一勝利 = False
      Call 結束遊戲
   Else '玩家二棄權
      玩家一勝利 = True
      Call 結束遊戲
   End If
   象棋框.Visible = False
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''設定
Private Sub 回主頁_Click()
   主頁框.Visible = True
   設定框.Visible = False
   Call 顯示按鈕
End Sub
Private Sub 炮跳開_Click()
   是否炮跳 = True
End Sub
Private Sub 炮跳關_Click()
   是否炮跳 = False
End Sub
Private Sub 音樂計時器_Timer()
   音樂計時器CountTime = 音樂計時器CountTime + 1
   If Temp2 = "0003" Then
      If 音樂計時器CountTime = 102 Then
         撥放器.Command = "CLOSE"
         音樂計時器CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "0006" Then
      If 音樂計時器CountTime = 73 Then
         撥放器.Command = "CLOSE"
         音樂計時器CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "Credits" Then
      If 音樂計時器CountTime = 205 Then
         撥放器.Command = "CLOSE"
         音樂計時器CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "Revolootin" Then
      If 音樂計時器CountTime = 72 Then
         撥放器.Command = "CLOSE"
         音樂計時器CountTime = 0
         Call 音樂播放
      End If
   ElseIf Temp2 = "SomeOfAKind" Then
      If 音樂計時器CountTime = 76 Then
         撥放器.Command = "CLOSE"
         音樂計時器CountTime = 0
         Call 音樂播放
      End If
   End If
End Sub
Private Sub 音樂開_Click()
   是否音樂 = True
   Call 音樂播放
End Sub
Private Sub 音樂關_Click()
   是否音樂 = False
   Call 音樂播放
End Sub
Public Sub 音樂播放()
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
   If 是否音樂 = True Then
      撥放器.FileName = App.Path & "\" & Temp2 & ".mp3"
      撥放器.Command = "OPEN"
      撥放器.Command = "PLAY"
   Else
      撥放器.Command = "CLOSE"
   End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''說明
Private Sub 說明確定_Click()
   主頁框.Visible = True
   說明框.Visible = False
   Call 顯示按鈕
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''關於
Private Sub 連結_Click()
   Shell "C:\Program Files\Internet Explorer\iexplore.exe & http://www.facebook.com/#!/groups/269973856388003/", vbMaximizedFocus
End Sub
Private Sub 連結_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   連結.ForeColor = vbRed
End Sub
Private Sub 連結_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   連結.ForeColor = &H80FF&
End Sub
Private Sub 確定_Click()
   主頁框.Visible = True
   關於框.Visible = False
   Call 顯示按鈕
End Sub
