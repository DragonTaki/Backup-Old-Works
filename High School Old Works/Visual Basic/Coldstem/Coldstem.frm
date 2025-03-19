VERSION 5.00
Begin VB.Form Jokes 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "Coldstem"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17250
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   17250
   StartUpPosition =   2  '螢幕中央
   Begin VB.Label ColdstemLabel 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      Caption         =   "冷梗研究社"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   36
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Jokes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Results As String
Private Sub ColdstemLabel_Click()
Dim Temp As Integer
   Call CheckFiles("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
   If Results = "" Then 'GoogleChrome Isn't Found On x64 computer
      Call CheckFiles("C:\Program Files\Google\Chrome\Application\chrome.exe")
      If Results = "" Then 'GoogleChrome Isn't Found On x86 computer
         Call CheckFiles("C:\Program Files\Internet Explorer\iexplore.exe")
         If Results = "" Then 'InternetExplorer Isn't Found
            Temp = MsgBox("Could Not Find Any Browser In This Computer!", vbOKOnly + vbExclamation, "Error")
         Else 'InternetExplorer Is Found
            Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.facebook.com/groups/coldstem/", vbMaximizedFocus
         End If
      Else 'GoogleChrome Is Found On x86 computer
         Shell "C:\Program Files\Google\Chrome\Application\chrome.exe http://www.facebook.com/groups/coldstem/", vbMaximizedFocus
      End If
   Else 'GoogleChrome Is Found On x64 computer
      Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe http://www.facebook.com/groups/coldstem/", vbMaximizedFocus
   End If
End Sub
Private Sub CheckFiles(ByRef TheFile As String)
   Results = Dir$(TheFile)
End Sub
