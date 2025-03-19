VERSION 5.00
Begin VB.Form MapleStory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '單線固定
   Caption         =   "MapleStory Installation Wizard"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   Icon            =   "MapleStory.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   3378.801
   StartUpPosition =   2  '螢幕中央
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Open All Link"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Beanfun Official"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "MapleStory Official"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "MapleStory Beanfun Start"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "webstartinst_tw_2_2_2_1.exe"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "webbeanfun_2_0_87_162_tw.exe"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "MapleStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'On Error Resume Next
Dim Results As String
Dim OpenAllLink As Boolean
Private Sub Label1_Click()
   Shell (App.Path & "\webbeanfun_2_0_87_162_tw.exe"), vbNormalFocus
End Sub
Private Sub Label2_Click()
   Shell (App.Path & "\webstartinst_tw_2_2_2_1.exe"), vbNormalFocus
End Sub
Private Sub Label3_Click()
   Call CheckFiles("C:\Program Files\Google\Chrome\Application\chrome.exe")
   If Results = "" Then 'GoogleChrome Isn't Found
      Call CheckFiles("C:\Program Files\Internet Explorer\iexplore.exe")
      If Results = "" Then 'InternetExplorer Isn't Found
         Temp = MsgBox("Could Not Find Any Browser In This Computer!", vbOKOnly + vbExclamation, "Error")
      Else 'InternetExplorer Is Found
         Shell "C:\Program Files\Internet Explorer\iexplore.exe http://tw.new.beanfun.com/game_zone/?service_code=610074&service_region=T9&sotp=876045", vbMaximizedFocus
      End If
   Else 'GoogleChrome Is Found
      Shell "C:\Program Files\Google\Chrome\Application\chrome.exe http://tw.new.beanfun.com/game_zone/?service_code=610074&service_region=T9&sotp=876045", vbMaximizedFocus
   End If
   If OpenAllLink = False Then
      End
   End If
End Sub
Private Sub Label4_Click()
   Call CheckFiles("C:\Program Files\Google\Chrome\Application\chrome.exe")
   If Results = "" Then 'GoogleChrome Isn't Found
      Call CheckFiles("C:\Program Files\Internet Explorer\iexplore.exe")
      If Results = "" Then 'InternetExplorer Isn't Found
         Temp = MsgBox("Could Not Find Any Browser In This Computer!", vbOKOnly + vbExclamation, "Error")
      Else 'InternetExplorer Is Found
         Shell "C:\Program Files\Internet Explorer\iexplore.exe http://tw.beanfun.com/maplestory/index.aspx", vbMaximizedFocus
      End If
   Else 'GoogleChrome Is Found
      Shell "C:\Program Files\Google\Chrome\Application\chrome.exe http://tw.beanfun.com/maplestory/index.aspx", vbMaximizedFocus
   End If
   If OpenAllLink = False Then
      End
   End If
End Sub
Private Sub Label5_Click()
   Call CheckFiles("C:\Program Files\Google\Chrome\Application\chrome.exe")
   If Results = "" Then 'GoogleChrome Isn't Found
      Call CheckFiles("C:\Program Files\Internet Explorer\iexplore.exe")
      If Results = "" Then 'InternetExplorer Isn't Found
         Temp = MsgBox("Could Not Find Any Browser In This Computer!", vbOKOnly + vbExclamation, "Error")
      Else 'InternetExplorer Is Found
         Shell "C:\Program Files\Internet Explorer\iexplore.exe http://tw.new.beanfun.com", vbMaximizedFocus
      End If
   Else 'GoogleChrome Is Found
      Shell "C:\Program Files\Google\Chrome\Application\chrome.exe http://tw.new.beanfun.com", vbMaximizedFocus
   End If
   If OpenAllLink = False Then
      End
   End If
End Sub
Private Sub Label6_Click()
   OpenAllLink = True
   Call Label3_Click
   Call Label4_Click
   Call Label5_Click
   End
End Sub
Private Sub CheckFiles(ByRef TheFile As String)
   Results = Dir$(TheFile)
End Sub
Private Sub FunctionNotOpen()
   Temp = MsgBox("This Function Has Not Be Opened!", vbOKOnly + vbExclamation, "Oops!")
End Sub
