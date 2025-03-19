Attribute VB_Name = "MakeFiles"
Private Sub Main()
   TheFile = App.Path & "\MakeFilesPath.txt"
   Results = Dir$(TheFile)
   If Results = "" Then '檔案不存在
      Temp = MsgBox("Cannot Find File 'MakeFilesPath.txt'.", vbExclamation, "Error")
   Else '檔案存在
      Open App.Path & "\MakeFilesPath.txt" For Input As #1 '讀檔
         Line Input #1, Temp
         If Dir(App.Path & "\" & Temp, vbDirectory) <> "" Then '檢查資料夾是否已存在
            For i = 2 To 10000
               If Dir(App.Path & "\" & Temp & " (" & i & ")", vbDirectory) <> "" Then '檢查下一個資料夾是否已存在
               '資料夾存在
               Else
                  MkDir (App.Path & "\" & Temp & " (" & i & ")") '資料夾不存在
                  End
               End If
            Next
         Else
            MkDir (App.Path & "\" & Temp) '資料夾不存在
         End If
      Close #1
   End If
   'MkDir (App.Path & "\ChineseChess")
End Sub
