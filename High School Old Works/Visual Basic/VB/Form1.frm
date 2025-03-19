VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ShowSpaceInfo(Drvpath)
  Dim fs, d, s
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(Drvpath)))
  s = "Drive " & d.DriveLetter & ":"
  s = s & vbCrLf
  s = s & "Total Size: " & FormatNumber(d.TotalSize / 1024, 0) & " Kbytes"
  s = s & "Available: " & FormatNumber(d.AvailableSpace / 1024, 0) & " Kbytes"
  MsgBox s
End Sub

Private Sub Form_Load()
Call ShowSpaceInfo("C")
End Sub
