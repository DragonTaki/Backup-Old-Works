VERSION 5.00
Begin VB.Form FileNameChanger 
   BackColor       =   &H00000000&
   Caption         =   "FileNameChanger"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "FileNameChanger.frx":0000
   ScaleHeight     =   1095
   ScaleWidth      =   6735
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Determine 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Determine"
      Height          =   1095
      Left            =   5880
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox NewFileName 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "Input New File Name"
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox FilePath 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Input File Path"
      Top             =   0
      Width           =   5895
   End
   Begin VB.TextBox OriginalFileName 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Input Original File Name"
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "FileNameChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Original_FilePath As String
Dim Original_OriginalFileName As String
Dim Original_NewFileName As String
Private Sub Determine_Click()
   If FilePath.Text = Original_FilePath And OriginalFileName.Text = Original_OriginalFileName And NewFileName.Text = Original_NewFileName Then 'User Hadn't Change Any Text
      Temp = MsgBox("You Have NOT Input Any Text!", vbExclamation, "Error")
   Else
      If FilePath.Text = "" And OriginalFileName.Text = "" And NewFileName.Text = "" Then 'User Hadn't Change Any Text
         Temp = MsgBox("You Have NOT Input Any Text!", vbExclamation, "Error")
      Else
         If FilePath.Text = "" Then 'File Path Is Empty
            Temp = MsgBox("The Path Cannot Be Empty!", vbExclamation, "Error")
         Else 'File Path Isn't Empty
            TheFile = FilePath.Text & "\" & OriginalFileName
            Results = Dir$(TheFile)
            If Results = "" Then 'The File Is NOT Exist
               Temp = MsgBox("The File Is NOT Exist!", vbExclamation, "Error")
            Else 'The File Is Exist
               If NewFileName.Text = "" Then 'The File New Name Is Enpty
                  Temp = MsgBox("The File New Name Cannot Be Enpty!", vbExclamation, "Error")
               Else 'The File New Name Isn't Enpty
                  ReName FilePath.Text, OriginalFileName.Text, NewFileName.Text
               End If
            End If
         End If
      End If
   End If
End Sub
Private Sub Form_Load()
   FileNameChanger.Width = 5000
   FileNameChanger.Height = 1600
   Original_FilePath = FilePath.Text
   Original_OriginalFileName = OriginalFileName.Text
   Original_NewFileName = NewFileName.Text
End Sub
Private Sub Form_Resize()
   If FileNameChanger.Height > FileNameChanger.Width - 1000 Then
      FileNameChanger.Height = FileNameChanger.Width - 1000
   End If
   '
   If FileNameChanger.Width < 5000 Then
      If FileNameChanger.WindowState = 0 Then
         FileNameChanger.Width = 5000
      End If
   ElseIf FileNameChanger.Height < 1600 Then
      If FileNameChanger.WindowState = 0 Then
         FileNameChanger.Height = 1600
      End If
   End If
   '
   Determine.Height = FileNameChanger.ScaleHeight
   Determine.Width = FileNameChanger.ScaleHeight
   Determine.Left = FileNameChanger.ScaleWidth - FileNameChanger.ScaleHeight
   '
   FilePath.Width = FileNameChanger.ScaleWidth - FileNameChanger.ScaleHeight
   OriginalFileName.Width = FileNameChanger.ScaleWidth - FileNameChanger.ScaleHeight
   NewFileName.Width = FileNameChanger.ScaleWidth - FileNameChanger.ScaleHeight
   '
   FilePath.Height = FileNameChanger.ScaleHeight / 3
   OriginalFileName.Height = FileNameChanger.ScaleHeight / 3
   NewFileName.Height = FileNameChanger.ScaleHeight / 3
   '
   FilePath.Top = 0
   OriginalFileName.Top = FilePath.Top + FilePath.Height
   NewFileName.Top = OriginalFileName.Top + OriginalFileName.Height
End Sub
Private Sub ReName(Folder As String, OldName As String, NewName As String)
      ChDir Folder
      Shell Environ("ComSpec") & " /c REN " & OldName & " " & NewName, vbHide
   End Sub
