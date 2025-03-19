VERSION 5.00
Begin VB.Form 象棋 
   BackColor       =   &H005FA76F&
   BorderStyle     =   1  '單線固定
   Caption         =   "象棋"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "M04_象棋頁.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13050
   StartUpPosition =   2  '螢幕中央
End
Attribute VB_Name = "象棋"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim 已按第一顆棋 As Boolean
Dim 已開局 As Boolean '選顏色用
Dim 棋已翻開(32) As Boolean '用位子記
Dim 空格(32) As Boolean '用位子記
Dim 現在換玩家二 As Boolean
Dim 玩家一是黑的 As Boolean
Dim 第一顆是自己的棋 As Boolean
Dim 第二顆是自己的棋 As Boolean
Dim 是可移動的 As Boolean
Dim 玩家一勝利 As Boolean
Dim 黑棋勝利 As Boolean
Dim 由右上角結束 As Boolean
Dim 已按返回 As Boolean
Dim 第一顆棋 As Integer
Dim 第二顆棋 As Integer
Dim 第一顆棋位置 As Integer
Dim 第二顆棋位置 As Integer
Dim 中間有棋 As Integer
Dim 結束圖片移動距離X As Integer '移動方向變數
Dim 結束圖片移動距離Y As Integer
Dim 結束訊息移動距離X As Integer
Dim 結束訊息移動距離Y As Integer
Private Sub 返回_Click()
   Ans = MsgBox("你確定要結束？", vbYesNo, "")
   If Ans = vbYes Then
      Name1P = Empty
      Name2P = Empty
      已按第一顆棋 = False
      已開局 = False
      For i = 0 To 32
         棋已翻開(i) = False
      Next
      For j = 0 To 32
         空格(j) = False
      Next
      現在換玩家二 = False
      玩家一是黑的 = False
      第一顆是自己的棋 = False
      第二顆是自己的棋 = False
      是可移動的 = False
      玩家一勝利 = False
      黑棋勝利 = False
      第一顆棋 = Empty
      第二顆棋 = Empty
      第一顆棋位置 = Empty
      第二顆棋位置 = Empty
      中間有棋 = Empty
      由右上角結束 = False
      已按返回 = True
      主頁.Show
      Unload 象棋
   Else
      Cancel = 1
   End If
End Sub
Private Sub 重新框_Click()
   結束圖片計時器.Enabled = False
   重新框.Visible = False
   Call 重新開始_Click
End Sub
Private Sub 結束圖片計時器_Timer()
   結束圖片.Left = 結束圖片.Left + 結束圖片移動距離X '移動X座標，物件的X標在<Left>這屬性內，其值越大代表越右方
   結束圖片.Top = 結束圖片.Top + 結束圖片移動距離Y '移動Y座標，物件的X標在<Top>這屬性內，其值越大，代表越下方
   If 結束圖片.Left + 結束圖片.Width > 象棋.Width Then '檢查是否超過表單右界(Form1.Width就是表單寬度)，但一個物件的左界值是<Left>，右界值是<Left>+<Width>
      結束圖片移動距離X = 0 - Abs(結束圖片移動距離X) '如果超過移動方向改成負值，代表須向左走
   End If
   If 結束圖片.Top + 結束圖片.Height > 象棋.Height - 600 Then '檢查是否超過表單下界(Form1.Height就是表單高度，代是含藍色標題區，所以要減600，使得到真的下界值)，
      結束圖片移動距離Y = 0 - Abs(結束圖片移動距離Y) '但一個物件的上界值是<Top>，下界值是<top>+<Height>,超過就把移動值改成負的，代表往上走
   End If 'Abs()取一個數的對值
   If 結束圖片.Left < 0 Then '物件一律相對表單移動，所以如果左界值<Left>小於0，代表撞到左牆了，就把值換成正的，改往右走
      結束圖片移動距離X = Abs(結束圖片移動距離X)
   End If
   If 結束圖片.Top < 0 Then '物件一律相對表單移動，所以如果上界值<Top>小於0，代表撞到上牆了，就把值換成正的，改往下走
      結束圖片移動距離Y = Abs(結束圖片移動距離Y)
   End If
   ''''''''''''''''''''''''''''''''''''''''
   結束訊息.Left = 結束訊息.Left + 結束訊息移動距離X '移動X座標，物件的X標在<Left>這屬性內，其值越大代表越右方
   結束訊息.Top = 結束訊息.Top + 結束訊息移動距離Y '移動Y座標，物件的X標在<Top>這屬性內，其值越大，代表越下方
   If 結束訊息.Left + 結束訊息.Width > 象棋.Width Then '檢查是否超過表單右界(Form1.Width就是表單寬度)，但一個物件的左界值是<Left>，右界值是<Left>+<Width>
      結束訊息移動距離X = 0 - Abs(結束訊息移動距離X) '如果超過移動方向改成負值，代表須向左走
   End If
   If 結束訊息.Top + 結束訊息.Height > 象棋.Height - 600 Then '檢查是否超過表單下界(Form1.Height就是表單高度，代是含藍色標題區，所以要減600，使得到真的下界值)，
      結束訊息移動距離Y = 0 - Abs(結束訊息移動距離Y) '但一個物件的上界值是<Top>，下界值是<top>+<Height>,超過就把移動值改成負的，代表往上走
   End If  'Abs()取一個數的對值
   If 結束訊息.Left < 0 Then '物件一律相對表單移動，所以如果左界值<Left>小於0，代表撞到左牆了，就把值換成正的，改往右走
      結束訊息移動距離X = Abs(結束訊息移動距離X)
   End If
   If 結束訊息.Top < 0 Then '物件一律相對表單移動，所以如果上界值<Top>小於0，代表撞到上牆了，就把值換成正的，改往下走
      結束訊息移動距離Y = Abs(結束訊息移動距離Y)
   End If
End Sub
Private Sub Form_Load()
   象棋框.Left = 0
   象棋框.Top = 0
   重新框.Left = 0
   重新框.Top = 0
   玩家一姓名.Caption = 主頁.Name1P
   玩家二姓名.Caption = 主頁.Name2P
   Randomize
   '生棋
   For i = 0 To 31
      If i > 0 Then
         Load 棋(i)
      End If
      棋(i).Left = 棋(0).Left + (i Mod 8) * (棋(0).Width + 240) '以0號位置為準，取8的餘數(代表要右偏幾格的寬度+間縫大小)
      棋(i).Top = 棋(0).Top + Int(i / 8) * (棋(0).Height + 105) '以0號位置為準，取整除8(代表要往下偏幾格的高度+間縫大小)
      棋(i).Visible = True
   Next
   Call 洗牌
   結束圖片移動距離X = 150 '把預定移動的速度在此決定，其在X代表左右，若為正，代表開始往右，若為負，代表起動時往左，其值越大，代表移動越快
   結束圖片移動距離Y = -100 '其在Y代表上下，若為正，代表開始往下，若為負，代表起動時往上，其值越大，代表移動越快
   結束訊息移動距離X = 150
   結束訊息移動距離Y = 100
End Sub
Private Sub 洗牌()
   Dim 棋已發過(32) As Boolean
   For i = 0 To 31
      棋已發過(i) = False
   Next
   For i = 0 To 31
重選棋:
      要放的棋 = Int(Rnd * 32)
      If 棋已發過(要放的棋) = True Then
         GoTo 重選棋
      End If
      棋已發過(要放的棋) = True
      棋(i).Tag = 要放的棋
   Next
End Sub
Private Sub 回到主頁_Click()
   Call Form_QueryUnload(0, 0)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If 已按返回 = True Then
      已按返回 = False
      Exit Sub
   End If
   Ans = MsgBox("你確定要結束？", vbYesNo, "")
      If Ans = vbYes Then
      Name1P = Empty
      Name2P = Empty
      已按第一顆棋 = False
      已開局 = False
      For i = 0 To 32
         棋已翻開(i) = False
      Next
      For j = 0 To 32
         空格(j) = False
      Next
      現在換玩家二 = False
      玩家一是黑的 = False
      第一顆是自己的棋 = False
      第二顆是自己的棋 = False
      是可移動的 = False
      玩家一勝利 = False
      黑棋勝利 = False
      第一顆棋 = Empty
      第二顆棋 = Empty
      第一顆棋位置 = Empty
      第二顆棋位置 = Empty
      中間有棋 = Empty
      由右上角結束 = False
      主頁.Show
      Unload 象棋
   Else
      Cancel = 1
   End If
End Sub
Private Sub 棋_Click(Index As Integer)
   If 已開局 = False Then '決定顏色
      If 16 > 棋(Index).Tag And 棋(Index).Tag >= 0 Then '點到黑棋
         玩家一.ForeColor = vbBlack
         玩家二.ForeColor = vbRed
         玩家一是黑的 = True
      Else '點到紅棋
         玩家一.ForeColor = vbRed
         玩家二.ForeColor = vbBlack
      End If
      Call 換人
      已開局 = True
      棋(Index).Picture = LoadPicture(App.Path & "\" & 棋(Index).Tag & ".che")
      棋已翻開(Index) = True
   Else '已決定顏色
      If 棋已翻開(Index) = False Then '該棋未翻開
         棋(Index).Picture = LoadPicture(App.Path & "\" & 棋(Index).Tag & ".che")
         棋已翻開(Index) = True
         Call 換人
      ElseIf 已按第一顆棋 = True And 棋已翻開(Index) = False Then '第二顆棋未翻開
         If 是否暗吃 = True Then '可暗吃
            GoTo 可暗吃 '視同已翻開
         Else '不可暗吃，離開
            Exit Sub
         End If
         '還沒按
      ElseIf 已按第一顆棋 = False Then '該棋已翻開，是第一顆
         已按第一顆棋 = True
         第一顆棋 = 棋(Index).Tag
         第一顆棋位置 = Index
         Call 判斷第一顆是否為自己的棋
         If 第一顆是自己的棋 = False Then '不是自己的棋
            已按第一顆棋 = False '還沒按
            第一顆棋 = Empty
            第一顆棋位置 = Empty
         Else '是自己的棋
            第一顆是自己的棋 = False
            持棋圖片.Picture = LoadPicture(App.Path & "\" & 棋(Index).Tag & ".che")
         End If
      ElseIf 已按第一顆棋 = True And 棋已翻開(Index) = True Then '該棋已翻開，是第二顆
         If 棋(Index).Tag = 第一顆棋 Then '和第一顆棋相同，取消持棋
            已按第一顆棋 = False '還沒按
            第一顆棋 = Empty
            第一顆棋位置 = Empty
            持棋圖片.Picture = LoadPicture(App.Path & "\空白.che")
         End If
         'If 第一顆棋 = 9 Or 第一顆棋 = 10 Then '例外：第一顆棋是包
         '   If 32 > 棋(Index).Tag And 棋(Index).Tag >= 27 Then '第二顆棋是兵
         '       Exit Sub '視同未按
         '   End If
         'ElseIf 第一顆棋 = 25 Or 第一顆棋 = 26 Then '例外：第一顆棋是炮
         '   If 16 > 棋(Index).Tag And 棋(Index).Tag >= 15 Then '第二顆棋是卒
         '       Exit Sub '視同未按
         '   End If
         'End If
可暗吃:
         第二顆棋 = 棋(Index).Tag
         第二顆棋位置 = Index
         Call 判斷第二顆是否為自己的棋
         If 第二顆是自己的棋 = True Then
            第二顆是自己的棋 = False
            第二顆棋 = Empty
            第二顆棋位置 = Empty
            Exit Sub '第二顆是自己的棋，不能吃，離開
         Else '第二顆不是自己的棋，可以吃
            Call 檢查兩顆棋的位置
            If 是可移動的 = False Then
               Exit Sub
            End If
            是可移動的 = False
            中間有棋 = 0
            Call 動棋子
         End If
      End If
   End If
   If Val(紅方剩棋.Caption) = 0 Then '結束遊戲，黑棋勝利
      If 玩家一是黑的 = True Then '玩家一勝利
         玩家一勝利 = True
         黑棋勝利 = True
         Call 結束遊戲
      Else '玩家二勝利
         玩家一勝利 = False
         黑棋勝利 = True
         Call 結束遊戲
      End If
   ElseIf Val(黑方剩棋.Caption) = 0 Then '結束遊戲，紅棋勝利
      If 玩家一是黑的 = False Then '玩家一勝利
         玩家一勝利 = True
         黑棋勝利 = False
         Call 結束遊戲
      Else '玩家二勝利
         玩家一勝利 = False
         黑棋勝利 = False
         Call 結束遊戲
      End If
   End If
End Sub
Private Sub 判斷第一顆是否為自己的棋()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 16 > 第一顆棋 And 第一顆棋 >= 0 Then '第一顆是黑棋
            第一顆是自己的棋 = True
         End If
      Else '玩家一是紅的
         If 32 > 第一顆棋 And 第一顆棋 >= 16 Then '第一顆是紅棋
            第一顆是自己的棋 = True
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 32 > 第一顆棋 And 第一顆棋 >= 16 Then '第一顆是紅棋
            第一顆是自己的棋 = True
         End If
      Else '玩家二是黑的
         If 16 > 第一顆棋 And 第一顆棋 >= 0 Then '第一顆是黑棋
            第一顆是自己的棋 = True
         End If
      End If
   End If
End Sub
Private Sub 判斷第二顆是否為自己的棋()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '第二顆是黑棋
            第二顆是自己的棋 = True
         End If
      Else '玩家一是紅的
         If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '第二顆是紅棋
            第二顆是自己的棋 = True
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '第二顆是紅棋
            第二顆是自己的棋 = True
         End If
      Else '玩家二是黑的
         If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '第二顆是黑棋
            第二顆是自己的棋 = True
         End If
      End If
   End If
End Sub
Private Sub 換人()
      If 現在換玩家二 = False Then
         現在換玩家二 = True
         換玩家一圖.Visible = False
         換玩家二圖.Visible = True
      Else
         現在換玩家二 = False
         換玩家一圖.Visible = True
         換玩家二圖.Visible = False
      End If
End Sub
Private Sub 檢查兩顆棋的位置()
   If 是否炮跳 = False Then '炮不跳
      GoTo 不開炮跳 '和一般棋動法一樣
   End If
   If 第一顆棋 = 9 Or 第一顆棋 = 10 Or 第一顆棋 = 25 Or 第一顆棋 = 26 Then  '第一顆棋是包或炮，在此移動
      If Abs((第一顆棋位置 - 第二顆棋位置)) < 8 Then '第一顆棋和第二顆棋同行
         If 棋(第二顆棋位置).Tag = -1 Then '第二顆棋是空的
            If 第一顆棋位置 = 第二顆棋位置 + 1 Or 第一顆棋位置 = 第二顆棋位置 - 1 Then '第二顆棋是空的且在第一顆棋上方或下方
               Call 動炮跳清除
            End If
         Else '第二顆棋不是空的
            Temp1 = 第一顆棋位置
            Temp2 = 第二顆棋位置
            If Temp1 > Temp2 Then
               Temp1 = 第二顆棋位置
               Temp2 = 第一顆棋位置
            End If
            For i = Temp1 + 1 To Temp2
               If 棋(i).Tag <> -1 Then
                  中間有棋 = 中間有棋 + 1
               End If
            Next
            If 中間有棋 = 2 Then
               Call 動炮跳清除
            End If
         End If
      ElseIf Abs(第一顆棋位置 - 第二顆棋位置) Mod 8 = 0 Then  '第一顆棋和第二顆棋同列
         If 棋(第二顆棋位置).Tag = -1 Then '第二顆棋是空的
            If 第一顆棋位置 = 第二顆棋位置 + 8 Or 第一顆棋位置 = 第二顆棋位置 - 8 Then '
               Call 動炮跳清除
            End If
         Else '第二顆棋不是空的
            Temp1 = 第一顆棋位置
            Temp2 = 第二顆棋位置
            If Temp1 > Temp2 Then
               Temp1 = 第二顆棋位置
               Temp2 = 第一顆棋位置
            End If
            For i = Temp1 + 8 To Temp2 Step 8
               If 棋(i).Tag <> -1 Then
                  中間有棋 = 中間有棋 + 1
               End If
            Next
            If 中間有棋 = 2 Then
               Call 動炮跳清除
            End If
         End If
      End If
   Else '第一顆棋不是包或炮
不開炮跳:
      If 第一顆棋位置 Mod 8 = 第二顆棋位置 Mod 8 And 第一顆棋位置 = 第二顆棋位置 + 8 Then '第二顆棋在第一顆棋上方
         是可移動的 = True
      ElseIf 第一顆棋位置 Mod 8 = 第二顆棋位置 Mod 8 And 第一顆棋位置 = 第二顆棋位置 - 8 Then '第二顆棋在第一顆棋下方
         是可移動的 = True
      ElseIf 第一顆棋位置 = 第二顆棋位置 + 1 Then '第二顆棋在第一顆棋右邊
         是可移動的 = True
      ElseIf 第一顆棋位置 = 第二顆棋位置 - 1 Then '第二顆棋在第一顆棋左邊
         是可移動的 = True
      End If
   End If
End Sub
Private Sub 吃完清除()
   已按第一顆棋 = False
   第一顆棋 = Empty
   第二顆棋 = Empty
   第一顆棋位置 = Empty
   第二顆棋位置 = Empty
   持棋圖片.Picture = LoadPicture(App.Path & "\空白.che")
   Call 換人
End Sub
Private Sub 動炮跳清除()
   If 現在換玩家二 = False Then '換玩家一
      If 玩家一是黑的 = True Then '玩家一是黑的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         End If
      Else '玩家一是紅的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         End If
      End If
   Else '換玩家二
      If 玩家一是黑的 = True Then '玩家二是紅的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         End If
      Else '玩家二是黑的
         If 第一顆棋位置 <> 第二顆棋位置 + 1 And 第一顆棋位置 <> 第二顆棋位置 - 1 And 第一顆棋位置 <> 第二顆棋位置 + 8 And 第一顆棋位置 <> 第二顆棋位置 - 8 Then
            紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         End If
      End If
   End If
   棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".che")
   棋(第一顆棋位置).Picture = LoadPicture("")
   持棋圖片.Picture = LoadPicture(App.Path & "\空白.che")
   Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
   空格(第一顆棋位置) = True
   第一顆棋 = Empty
   第二顆棋 = Empty
   第一顆棋位置 = Empty
   第二顆棋位置 = Empty
   已按第一顆棋 = False
   中間有棋 = 0
   Call 換人
End Sub
Private Sub 動棋子()
   If 現在換玩家二 = False Then
      If 玩家一是黑的 = True Then
         Call 動黑棋
      Else '玩家一是紅的
         Call 動紅棋
      End If
   Else    '現在換玩家二
      If 玩家一是黑的 = True Then
         Call 動紅棋
      Else '玩家二是黑的
         Call 動黑棋
      End If
   End If
End Sub
Private Sub 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag()
   棋(第二顆棋位置).Tag = 棋(第一顆棋位置).Tag
   棋(第一顆棋位置).Tag = -1
End Sub
Private Sub 移棋子一至棋子二()
   棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".che")
   棋(第一顆棋位置).Picture = LoadPicture("")
End Sub
Private Sub 動黑棋()
   If 第一顆棋 = 0 Then '將
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 27 > 第二顆棋 And 第二顆棋 >= 16 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 27 Then '將吃兵
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 1 Or 第一顆棋 = 2 Then '士
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 17 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 16 Then '士吃帥
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 3 Or 第一顆棋 = 4 Then '象
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 19 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 17 Or 第二顆棋 = 18 Then '象吃帥仕
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 5 Or 第一顆棋 = 6 Then '車
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 21 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 21 > 第二顆棋 And 第二顆棋 >= 16 Then '車吃帥仕相
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 7 Or 第一顆棋 = 8 Then '馬
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 23 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 23 > 第二顆棋 And 第二顆棋 >= 16 Then '馬吃帥仕相硨
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 9 Or 第一顆棋 = 10 Then '包
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 32 > 第二顆棋 And 第二顆棋 >= 27 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 27 > 第二顆棋 And 第二顆棋 >= 16 Then
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 16 > 第一顆棋 And 第一顆棋 >= 11 Then '卒
      If 16 > 第二顆棋 And 第二顆棋 >= 0 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 25 Or 第二顆棋 = 26 Then '吃炮
         '不動
      ElseIf 第二顆棋 = 16 Or 第二顆棋 = 27 Or 第二顆棋 = 28 Or 第二顆棋 = 29 Or 第二顆棋 = 30 Or 第二顆棋 = 31 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 25 > 第二顆棋 And 第二顆棋 >= 17 Then    '卒吃仕相硨碼
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   End If
End Sub
Private Sub 動紅棋()
   If 第一顆棋 = 16 Then '帥
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 11 > 第二顆棋 And 第二顆棋 >= 0 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 15 > 第二顆棋 And 第二顆棋 > 11 Then '帥吃卒
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 17 Or 第一顆棋 = 18 Then '仕
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then   '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Then '仕吃將
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 19 Or 第一顆棋 = 20 Then '相
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Then '相吃將士
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 21 Or 第一顆棋 = 22 Then '硨
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Then '相吃將士象
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 23 Or 第一顆棋 = 24 Then '碼
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then  '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Then '相吃將士象車
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 第一顆棋 = 25 Or 第一顆棋 = 26 Then '炮
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃異色
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 9 Or 第二顆棋 = 10 Then
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   ElseIf 32 > 第一顆棋 And 第一顆棋 >= 27 Then '兵
      If 32 > 第二顆棋 And 第二顆棋 >= 16 Then '吃同色
         '自己棋，不動
      ElseIf 第二顆棋 = 9 Or 第二顆棋 = 10 Then '吃包
         '不動
      ElseIf 第二顆棋 = 0 Or 第二顆棋 = 7 Or 第二顆棋 = 8 Or 第二顆棋 = 11 Or 第二顆棋 = 12 Or 第二顆棋 = 13 Or 第二顆棋 = 14 Or 第二顆棋 = 15 Then '吃異色
         棋(第二顆棋位置).Picture = LoadPicture(App.Path & "\" & 第一顆棋 & ".che")
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         黑方剩棋.Caption = Val(黑方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 第二顆棋 = 1 Or 第二顆棋 = 2 Or 第二顆棋 = 3 Or 第二顆棋 = 4 Or 第二顆棋 = 5 Or 第二顆棋 = 6 Then '兵吃士象車馬
         棋(第一顆棋位置).Picture = LoadPicture("")
         空格(第一顆棋位置) = True
         棋(第一顆棋位置).Tag = -1
         紅方剩棋.Caption = Val(紅方剩棋.Caption) - 1
         Call 吃完清除
      ElseIf 空格(第二顆棋位置) = True Then
         Call 移棋子一至棋子二
         空格(第一顆棋位置) = True
         空格(第二顆棋位置) = False
         Call 第二顆棋的Tag換第一顆棋的Tag並清掉第一顆棋的Tag
         Call 吃完清除
      End If
   End If
End Sub
Private Sub 結束遊戲()
   結束圖片計時器.Enabled = True
   重新框.Visible = True
   If 玩家一勝利 = True Then
      結束訊息.Caption = Name1P & "勝利"
   Else
      結束訊息.Caption = Name2P & "勝利"
   End If
End Sub
Private Sub 重新開始_Click()
   Call 洗牌
   For i = 0 To 31 '蓋牌
      棋(i).Picture = LoadPicture(App.Path & "\空白.che")
   Next
   黑方剩棋.Caption = "16"
   紅方剩棋.Caption = "16"
   換玩家一圖.Visible = True
   換玩家二圖.Visible = False
   現在換玩家二 = False
   已按第一顆棋 = False
   For i = 0 To 32
      棋已翻開(i) = False
   Next
   For j = 0 To 32
      空格(j) = False
   Next
   已開局 = False
   現在換玩家二 = False
   玩家一是黑的 = False
   第一顆是自己的棋 = False
   第二顆是自己的棋 = False
   是可移動的 = False
   玩家一勝利 = False
   黑棋勝利 = False
   第一顆棋 = Empty
   第二顆棋 = Empty
   第一顆棋位置 = Empty
   第二顆棋位置 = Empty
   中間有棋 = Empty
   持棋圖片.Picture = LoadPicture(App.Path & "\空白.che")
End Sub
Private Sub 棄權_Click()
   If 現在換玩家二 = False Then '玩家一棄權
      玩家一勝利 = False
      Call 結束遊戲
   Else '玩家二棄權
      玩家一勝利 = True
      Call 結束遊戲
   End If
End Sub
