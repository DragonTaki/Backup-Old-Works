Attribute VB_Name = "MakeFiles"
Private Sub Main()
   TheFile = App.Path & "\MakeFilesPath.txt"
   Results = Dir$(TheFile)
   If Results = "" Then '�ɮפ��s�b
      Temp = MsgBox("Cannot Find File 'MakeFilesPath.txt'.", vbExclamation, "Error")
   Else '�ɮצs�b
      Open App.Path & "\MakeFilesPath.txt" For Input As #1 'Ū��
         Line Input #1, Temp
         If Dir(App.Path & "\" & Temp, vbDirectory) <> "" Then '�ˬd��Ƨ��O�_�w�s�b
            For i = 2 To 10000
               If Dir(App.Path & "\" & Temp & " (" & i & ")", vbDirectory) <> "" Then '�ˬd�U�@�Ӹ�Ƨ��O�_�w�s�b
               '��Ƨ��s�b
               Else
                  MkDir (App.Path & "\" & Temp & " (" & i & ")") '��Ƨ����s�b
                  End
               End If
            Next
         Else
            MkDir (App.Path & "\" & Temp) '��Ƨ����s�b
         End If
      Close #1
   End If
   'MkDir (App.Path & "\ChineseChess")
End Sub
