VERSION 5.00
Begin VB.Form TestExecution 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "TestExecution"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "TestExecution.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Try 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Try"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00000000&
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.DriveListBox Drive 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Calculate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calculate"
      Height          =   375
      Left            =   5640
      MaskColor       =   &H00000000&
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer InstallTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6480
      Top             =   360
   End
   Begin VB.CommandButton Install 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Install"
      Height          =   375
      Left            =   5640
      MaskColor       =   &H00000000&
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton OpenFileButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go"
      Height          =   375
      Left            =   6480
      MaskColor       =   &H00000000&
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox OpenFile 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Open File ......"
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label DriveLabel 
      BackColor       =   &H00000000&
      Caption         =   "Calculate Drive Bytes"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label 第二安裝頁框正在安裝檔案名稱 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   5655
   End
End
Attribute VB_Name = "TestExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InstallCount As Long
Dim InstallNewPath As String
Dim InstallNewFilePath As String
Dim InstallOriginalPath As String
Dim CanDoNext As Boolean
Dim ObjPrc As Object
Dim StrSQL As String
Private Sub Calculate_Click()
   Call ShowSpaceInfo(Drive.Drive)
End Sub
Sub ShowSpaceInfo(Drvpath)
   Dim fs, d, Sum
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(Drvpath)))
   Sum = "Drive " & d.DriveLetter & ":"
   Sum = Sum & vbCrLf
   Sum = Sum & "  Total Size: " & FormatNumber(d.TotalSize / 1024, 0) & " Kbytes"
   Sum = Sum & vbCrLf
   Sum = Sum & "  Available: " & FormatNumber(d.AvailableSpace / 1024, 0) & " Kbytes"
   Temp = MsgBox(Sum, vbInformation, "")
End Sub
Private Sub Drive_Change()
   DriveCount = Drive.Drive
End Sub
Private Sub Form_Load() '測試區
'Call TestExeRunning
'End
End Sub
Private Sub Install_Click()
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
      Call Kill_file("C:\Program Files\MakeFilesPath.txt")
      Call Kill_file("C:\Program Files\FilePath.dll")
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
      Call Kill_file("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_file("C:\Program Files\ChineseChess\FilePath.dll")
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
      Call Kill_file("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_file("C:\Program Files\ChineseChess\FilePath.dll")
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
      Call Kill_file("C:\Program Files\ChineseChess\MakeFilesPath.txt")
      Call Kill_file("C:\Program Files\ChineseChess\FilePath.dll")
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
   For i = 1 To 100
      For j = 1 To 100
      CountTime = CountTime + 1
      Next
   Next
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
   '
   End
End Sub
Private Sub Kill_file(ByRef Dir_string As String)
On Error GoTo Exit_out
Kill Dir_string
Exit_out:
End Sub
Private Sub RealInstallCount() '真正安裝內容
   '
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
Private Sub InstallTimer_Timer()
   CanDoNext = True
   InstallTimer.Enabled = False
End Sub
Private Sub OpenFileButton_Click()
   Results = Dir$(OpenFile.Text)
   If Results = "" Then '路徑不存在，嘗試加上此檔所在路徑
      Results = Dir$(App.Path & "\" & OpenFile.Text)
      If Results = "" Then '路徑不存在
         Temp = MsgBox("Cannot Find This File!", vbExclamation, "Error")
         OpenFile.Text = ""
      Else '加上此檔所在路徑檔案存在
         Temp = Shell(App.Path & "\" & OpenFile.Text, vbNormalFocus)
      End If
   Else '路徑存在
      If OpenFile.Text = "" Then '路徑空白
         Temp = MsgBox("The Path Cannot Be Empty!", vbExclamation, "Error")
      Else
         Shell (OpenFile.Text), vbNormalFocus
      End If
   End If
End Sub
Private Sub Try_Click()
   Tmp1 = Time
   For i = 1 To 10000
      For j = 1 To 1000
      CountTime = CountTime + 1
      Next
   Next
   Tmp2 = Time
   Temp = MsgBox(Tmp2 - Tmp1, vbInformation, "Total Time")
End Sub
Private Sub TestExeRunning()
   StrSQL = "Select * from Win32_Process Where Name = 'TestExecution 1.0.6.exe'"
   If GetObject("winmgmts:").ExecQuery(StrSQL).Count > 0 Then
      'Temp = MsgBox("程式已運行", vbInformation, "")
   Else
      'Temp = MsgBox("程式未運行", vbInformation, "")
   End If
End Sub
