Attribute VB_Name = "Module1"
Private Sub Main()
   If Dir("c:\Program Files", vbDirectory) <> "" Then
      MsgBox "資料夾存在"
   Else
      MsgBox "資料夾不存在"
   End If
End Sub
'http://www.blueshop.com.tw/board/show.asp?subcde=BRD20060630090126N7I&fumcde=FUM200501271723350KG

