VERSION 5.00
Begin VB.Form HideBorderStyle 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "HideBorderStyle"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12045
   ForeColor       =   &H00FFFFFF&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   12045
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox EndText 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "End"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox HideBorderStyleText 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "HideBorderStyle.frx":0000
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "HideBorderStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_SYSMENU = &H80000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Const GWL_STYLE = (-16)
Private Sub EndText_Click()
   Clipboard.Clear
   Clipboard.SetText HideBorderStyleText.Text
   End
End Sub
Private Sub Form_DblClick()
   Unload Me
End Sub
Private Sub Form_Load()
   Dim TempLng As Long
   TempLng = GetWindowLong(Me.hwnd, GWL_STYLE)
   TempLng = TempLng And Not WS_SYSMENU
   TempLng = TempLng And Not WS_MINIMIZEBOX
   TempLng = TempLng And Not WS_MAXIMIZEBOX
   SetWindowLong Me.hwnd, GWL_STYLE, TempLng
   HideBorderStyle.Height = HideBorderStyleText.Height + 600
   HideBorderStyle.Width = HideBorderStyleText.Width
End Sub
