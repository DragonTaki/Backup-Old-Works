VERSION 5.00
Begin VB.Form MsgboxForm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "Msgbox"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Msgbox.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6495
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame MsgboxPictureFrame 
      BackColor       =   &H00000000&
      Caption         =   "Msgbox Picture"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   14
      Top             =   240
      Width           =   1935
      Begin VB.OptionButton MsgboxPicture 
         BackColor       =   &H00000000&
         Caption         =   "Critical"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton MsgboxPicture 
         BackColor       =   &H00000000&
         Caption         =   "Question"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton MsgboxPicture 
         BackColor       =   &H00000000&
         Caption         =   "Information"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   64
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton MsgboxPicture 
         BackColor       =   &H00000000&
         Caption         =   "Exclamation"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   48
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton MsgboxPicture 
         BackColor       =   &H00000000&
         Caption         =   "None"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox MsgboxTitleTextBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox MultipleSumNumber 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   0
      Width           =   975
   End
   Begin VB.OptionButton Multiple 
      BackColor       =   &H00000000&
      Caption         =   "Singular"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.Frame MsgboxButtonFrame 
      BackColor       =   &H00000000&
      Caption         =   "Msgbox Button"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
      Begin VB.OptionButton MsgboxButton 
         BackColor       =   &H00000000&
         Caption         =   "RetryCancel"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton MsgboxButton 
         BackColor       =   &H00000000&
         Caption         =   "YesNo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton MsgboxButton 
         BackColor       =   &H00000000&
         Caption         =   "OKOnly"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton MsgboxButton 
         BackColor       =   &H00000000&
         Caption         =   "AbortRetryIgnore"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton MsgboxButton 
         BackColor       =   &H00000000&
         Caption         =   "OKCanael"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton MsgboxButton 
         BackColor       =   &H00000000&
         Caption         =   "YesNoCencel"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.TextBox MultipleNowNumber 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox MsgboxTextTextBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   4575
   End
   Begin VB.OptionButton Multiple 
      BackColor       =   &H00000000&
      Caption         =   "Multiple"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   17
      Top             =   0
      Width           =   975
   End
   Begin VB.Label NextSum 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Next"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label MakeMsgbox 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Make Msgbox!"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label PreviousSum 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Previous"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "MsgboxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgboxPictureValue As Integer
Dim MsgboxButtonValue As Integer
Dim MsgboxTextTextBoxText As String
Dim MsgboxTitleTextBoxText As String
Dim TempTitle As String
Dim TempText As String
Const OriginalTitle = "Msgbox Title"
Const OriginalText = "Msgbox Text"
Private Sub Form_Load()
   App.TaskVisible = False
   'MsgboxForm.ShowIcon = False
   MsgboxTitleTextBox.Text = OriginalTitle
   MsgboxTextTextBox.Text = OriginalText
End Sub
Private Sub MakeMsgbox_Click()
   If MsgboxPicture(0).Value = True Then
      MsgboxPictureValue = 0
   ElseIf MsgboxPicture(16).Value = True Then
      MsgboxPictureValue = 16
   ElseIf MsgboxPicture(32).Value = True Then
      MsgboxPictureValue = 32
   ElseIf MsgboxPicture(48).Value = True Then
      MsgboxPictureValue = 48
   ElseIf MsgboxPicture(64).Value = True Then
      MsgboxPictureValue = 64
   End If
   If MsgboxButton(0).Value = True Then
      MsgboxButtonValue = 0
   ElseIf MsgboxButton(1).Value = True Then
      MsgboxButtonValue = 1
   ElseIf MsgboxButton(2).Value = True Then
      MsgboxButtonValue = 2
   ElseIf MsgboxButton(3).Value = True Then
      MsgboxButtonValue = 3
   ElseIf MsgboxButton(4).Value = True Then
      MsgboxButtonValue = 4
   ElseIf MsgboxButton(5).Value = True Then
      MsgboxButtonValue = 5
   End If
   MsgboxTextTextBoxText = MsgboxTextTextBox.Text
   MsgboxTitleTextBoxText = MsgboxTitleTextBox.Text
   If MsgboxTitleTextBoxText = OriginalTitle And MsgboxTextTextBoxText = OriginalText Then
      TempTitle = "Shit!"
      TempText = "You are the son of the bitch!" & vbCrLf & vbCrLf & "Don't you know you should input the title and the text?"
   Else
      TempTitle = MsgboxTitleTextBoxText
      TempText = MsgboxTextTextBoxText
   End If
   Unload MsgboxForm
   Temp = MsgBox(TempText, MsgboxPictureValue + MsgboxButtonValue, TempTitle)
   Load MsgboxForm
   MsgboxForm.Show
   MsgboxTextTextBox.Text = MsgboxTextTextBoxText
   MsgboxTitleTextBox.Text = MsgboxTitleTextBoxText
   MsgboxPicture(MsgboxPictureValue).Value = True
   MsgboxButton(MsgboxButtonValue).Value = True
End Sub
Private Sub Multiple_Click(Index As Integer)
   If Multiple(0).Value = True Then
      MultipleSumNumber.Enabled = False
      MultipleNowNumber.Enabled = False
      MultipleSumNumber.Text = ""
      MultipleNowNumber.Text = ""
   Else
      MultipleSumNumber.Enabled = True
      MultipleNowNumber.Enabled = True
      MultipleSumNumber.Text = "1"
      MultipleNowNumber.Text = "1"
   End If
End Sub
Private Sub MultipleSumNumber_Change()
   Dim MagboxSum(10000) As Integer
   
End Sub
