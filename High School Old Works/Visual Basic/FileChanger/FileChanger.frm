VERSION 5.00
Begin VB.Form FileChanger 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "FileChanger"
   ClientHeight    =   3315
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8370
   ForeColor       =   &H00000000&
   Icon            =   "FileChanger.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8370
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame FilePropertyFrame 
      BackColor       =   &H00000000&
      Caption         =   "File Property"
      ForeColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton FileProperty 
         BackColor       =   &H00000000&
         Caption         =   "Archive"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton FileProperty 
         BackColor       =   &H00000000&
         Caption         =   "Normal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton FileProperty 
         BackColor       =   &H00000000&
         Caption         =   "System"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton FileProperty 
         BackColor       =   &H00000000&
         Caption         =   "Volume"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton FileProperty 
         BackColor       =   &H00000000&
         Caption         =   "Hidden"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton FileProperty 
         BackColor       =   &H00000000&
         Caption         =   "ReadOnly"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.DriveListBox FileDrive 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   60
      Width           =   3375
   End
   Begin VB.DirListBox FileDir 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   1320
      TabIndex        =   6
      Top             =   360
      Width           =   3375
   End
   Begin VB.FileListBox FileFile 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2250
      Hidden          =   -1  'True
      Left            =   4680
      TabIndex        =   9
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label ChangeButton 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "Change"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "FileChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Shared Function GetFolderPath ( folder As Environment.SpecialFolder ) As String
Private Sub ChangeButton_Click()
   TempPath = (FileFile.Path & "\" & FileFile.FileName)
   On Error GoTo FileSetAttrError
   If FileProperty(0).Value = True Then 'Normal'0
      SetAttr (TempPath), vbNormal
   ElseIf FileProperty(1).Value = True Then 'ReadOnly'1
      SetAttr (TempPath), vbReadOnly
   ElseIf FileProperty(2).Value = True Then 'Hidden'2
      SetAttr (TempPath), vbHidden
   ElseIf FileProperty(4).Value = True Then 'System'4
      SetAttr (TempPath), vbSystem
   ElseIf FileProperty(8).Value = True Then 'Volume'8
      SetAttr (TempPath), vbVolume
   ElseIf FileProperty(32).Value = True Then 'Archive'32
      SetAttr (TempPath), vbArchive
   End If
   '0=Normal 1=ReadOnly 2=Hidden 4=System 8=Volume 16=Directory 32=Archive 64=Alias
   Temp = MsgBox("The file has changed!", vbInformation + vbOKOnly, "")
   FileProperty(0).Value = True
   Exit Sub
FileSetAttrError:
   Temp = MsgBox("This File Can't Set!", vbOKOnly + vbExclamation, "Drive Error")
End Sub
Private Sub FileDir_Change()
   FileFile.Path = FileDir.Path
   Call DirList(FileDir.Path)
End Sub
Private Sub FileDrive_Change()
   On Error GoTo DriveError
   FileDir.Path = FileDrive.Drive
   Exit Sub
DriveError:
   Temp = MsgBox("This Drive Can't Read!", vbOKOnly + vbExclamation, "Drive Error")
End Sub
Public Sub DirList(Path As String)
   Dim fsobj, Fold, d, fs, s
   Set fsobj = CreateObject("Scripting.FileSystemObject")
   Set Fold = fsobj.GetFolder(Path)
   Set d = Fold.SubFolders
   'List1.Clear
   'List2.Clear
   'For i = 0 To FileDir.ListCount - 1
   '   List2.AddItem FileDir.List(i)
   'Next
   'For Each fs In d
   'List1.AddItem fs.Name
   'List2.AddItem fs.Name
   'Next
   
End Sub

