Attribute VB_Name = "Module1"
Private Sub Main()
   If Dir("c:\Program Files", vbDirectory) <> "" Then
      MsgBox "��Ƨ��s�b"
   Else
      MsgBox "��Ƨ����s�b"
   End If
End Sub
'http://www.blueshop.com.tw/board/show.asp?subcde=BRD20060630090126N7I&fumcde=FUM200501271723350KG

