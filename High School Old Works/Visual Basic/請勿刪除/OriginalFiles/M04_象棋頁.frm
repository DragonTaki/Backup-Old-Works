VERSION 5.00
Begin VB.Form �H�� 
   BackColor       =   &H005FA76F&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�H��"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13050
   Icon            =   "M04_�H�ѭ�.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13050
   StartUpPosition =   2  '�ù�����
End
Attribute VB_Name = "�H��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim �w���Ĥ@���� As Boolean
Dim �w�}�� As Boolean '���C���
Dim �Ѥw½�}(32) As Boolean '�Φ�l�O
Dim �Ů�(32) As Boolean '�Φ�l�O
Dim �{�b�����a�G As Boolean
Dim ���a�@�O�ª� As Boolean
Dim �Ĥ@���O�ۤv���� As Boolean
Dim �ĤG���O�ۤv���� As Boolean
Dim �O�i���ʪ� As Boolean
Dim ���a�@�ӧQ As Boolean
Dim �´ѳӧQ As Boolean
Dim �ѥk�W������ As Boolean
Dim �w����^ As Boolean
Dim �Ĥ@���� As Integer
Dim �ĤG���� As Integer
Dim �Ĥ@���Ѧ�m As Integer
Dim �ĤG���Ѧ�m As Integer
Dim �������� As Integer
Dim �����Ϥ����ʶZ��X As Integer '���ʤ�V�ܼ�
Dim �����Ϥ����ʶZ��Y As Integer
Dim �����T�����ʶZ��X As Integer
Dim �����T�����ʶZ��Y As Integer
Private Sub ��^_Click()
   Ans = MsgBox("�A�T�w�n�����H", vbYesNo, "")
   If Ans = vbYes Then
      Name1P = Empty
      Name2P = Empty
      �w���Ĥ@���� = False
      �w�}�� = False
      For i = 0 To 32
         �Ѥw½�}(i) = False
      Next
      For j = 0 To 32
         �Ů�(j) = False
      Next
      �{�b�����a�G = False
      ���a�@�O�ª� = False
      �Ĥ@���O�ۤv���� = False
      �ĤG���O�ۤv���� = False
      �O�i���ʪ� = False
      ���a�@�ӧQ = False
      �´ѳӧQ = False
      �Ĥ@���� = Empty
      �ĤG���� = Empty
      �Ĥ@���Ѧ�m = Empty
      �ĤG���Ѧ�m = Empty
      �������� = Empty
      �ѥk�W������ = False
      �w����^ = True
      �D��.Show
      Unload �H��
   Else
      Cancel = 1
   End If
End Sub
Private Sub ���s��_Click()
   �����Ϥ��p�ɾ�.Enabled = False
   ���s��.Visible = False
   Call ���s�}�l_Click
End Sub
Private Sub �����Ϥ��p�ɾ�_Timer()
   �����Ϥ�.Left = �����Ϥ�.Left + �����Ϥ����ʶZ��X '����X�y�СA����X�Цb<Left>�o�ݩʤ��A��ȶV�j�N��V�k��
   �����Ϥ�.Top = �����Ϥ�.Top + �����Ϥ����ʶZ��Y '����Y�y�СA����X�Цb<Top>�o�ݩʤ��A��ȶV�j�A�N��V�U��
   If �����Ϥ�.Left + �����Ϥ�.Width > �H��.Width Then '�ˬd�O�_�W�L���k��(Form1.Width�N�O���e��)�A���@�Ӫ��󪺥��ɭȬO<Left>�A�k�ɭȬO<Left>+<Width>
      �����Ϥ����ʶZ��X = 0 - Abs(�����Ϥ����ʶZ��X) '�p�G�W�L���ʤ�V�令�t�ȡA�N���V����
   End If
   If �����Ϥ�.Top + �����Ϥ�.Height > �H��.Height - 600 Then '�ˬd�O�_�W�L���U��(Form1.Height�N�O��氪�סA�N�O�t�Ŧ���D�ϡA�ҥH�n��600�A�ϱo��u���U�ɭ�)�A
      �����Ϥ����ʶZ��Y = 0 - Abs(�����Ϥ����ʶZ��Y) '���@�Ӫ��󪺤W�ɭȬO<Top>�A�U�ɭȬO<top>+<Height>,�W�L�N�Ⲿ�ʭȧ令�t���A�N���W��
   End If 'Abs()���@�Ӽƪ����
   If �����Ϥ�.Left < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G���ɭ�<Left>�p��0�A�N���쥪��F�A�N��ȴ��������A�啕�k��
      �����Ϥ����ʶZ��X = Abs(�����Ϥ����ʶZ��X)
   End If
   If �����Ϥ�.Top < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G�W�ɭ�<Top>�p��0�A�N����W��F�A�N��ȴ��������A�啕�U��
      �����Ϥ����ʶZ��Y = Abs(�����Ϥ����ʶZ��Y)
   End If
   ''''''''''''''''''''''''''''''''''''''''
   �����T��.Left = �����T��.Left + �����T�����ʶZ��X '����X�y�СA����X�Цb<Left>�o�ݩʤ��A��ȶV�j�N��V�k��
   �����T��.Top = �����T��.Top + �����T�����ʶZ��Y '����Y�y�СA����X�Цb<Top>�o�ݩʤ��A��ȶV�j�A�N��V�U��
   If �����T��.Left + �����T��.Width > �H��.Width Then '�ˬd�O�_�W�L���k��(Form1.Width�N�O���e��)�A���@�Ӫ��󪺥��ɭȬO<Left>�A�k�ɭȬO<Left>+<Width>
      �����T�����ʶZ��X = 0 - Abs(�����T�����ʶZ��X) '�p�G�W�L���ʤ�V�令�t�ȡA�N���V����
   End If
   If �����T��.Top + �����T��.Height > �H��.Height - 600 Then '�ˬd�O�_�W�L���U��(Form1.Height�N�O��氪�סA�N�O�t�Ŧ���D�ϡA�ҥH�n��600�A�ϱo��u���U�ɭ�)�A
      �����T�����ʶZ��Y = 0 - Abs(�����T�����ʶZ��Y) '���@�Ӫ��󪺤W�ɭȬO<Top>�A�U�ɭȬO<top>+<Height>,�W�L�N�Ⲿ�ʭȧ令�t���A�N���W��
   End If  'Abs()���@�Ӽƪ����
   If �����T��.Left < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G���ɭ�<Left>�p��0�A�N���쥪��F�A�N��ȴ��������A�啕�k��
      �����T�����ʶZ��X = Abs(�����T�����ʶZ��X)
   End If
   If �����T��.Top < 0 Then '����@�߬۹��沾�ʡA�ҥH�p�G�W�ɭ�<Top>�p��0�A�N����W��F�A�N��ȴ��������A�啕�U��
      �����T�����ʶZ��Y = Abs(�����T�����ʶZ��Y)
   End If
End Sub
Private Sub Form_Load()
   �H�Ѯ�.Left = 0
   �H�Ѯ�.Top = 0
   ���s��.Left = 0
   ���s��.Top = 0
   ���a�@�m�W.Caption = �D��.Name1P
   ���a�G�m�W.Caption = �D��.Name2P
   Randomize
   '�ʹ�
   For i = 0 To 31
      If i > 0 Then
         Load ��(i)
      End If
      ��(i).Left = ��(0).Left + (i Mod 8) * (��(0).Width + 240) '�H0����m���ǡA��8���l��(�N��n�k���X�檺�e��+���_�j�p)
      ��(i).Top = ��(0).Top + Int(i / 8) * (��(0).Height + 105) '�H0����m���ǡA���㰣8(�N��n���U���X�檺����+���_�j�p)
      ��(i).Visible = True
   Next
   Call �~�P
   �����Ϥ����ʶZ��X = 150 '��w�w���ʪ��t�צb���M�w�A��bX�N���k�A�Y�����A�N��}�l���k�A�Y���t�A�N��_�ʮɩ����A��ȶV�j�A�N���ʶV��
   �����Ϥ����ʶZ��Y = -100 '��bY�N��W�U�A�Y�����A�N��}�l���U�A�Y���t�A�N��_�ʮɩ��W�A��ȶV�j�A�N���ʶV��
   �����T�����ʶZ��X = 150
   �����T�����ʶZ��Y = 100
End Sub
Private Sub �~�P()
   Dim �Ѥw�o�L(32) As Boolean
   For i = 0 To 31
      �Ѥw�o�L(i) = False
   Next
   For i = 0 To 31
�����:
      �n�񪺴� = Int(Rnd * 32)
      If �Ѥw�o�L(�n�񪺴�) = True Then
         GoTo �����
      End If
      �Ѥw�o�L(�n�񪺴�) = True
      ��(i).Tag = �n�񪺴�
   Next
End Sub
Private Sub �^��D��_Click()
   Call Form_QueryUnload(0, 0)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If �w����^ = True Then
      �w����^ = False
      Exit Sub
   End If
   Ans = MsgBox("�A�T�w�n�����H", vbYesNo, "")
      If Ans = vbYes Then
      Name1P = Empty
      Name2P = Empty
      �w���Ĥ@���� = False
      �w�}�� = False
      For i = 0 To 32
         �Ѥw½�}(i) = False
      Next
      For j = 0 To 32
         �Ů�(j) = False
      Next
      �{�b�����a�G = False
      ���a�@�O�ª� = False
      �Ĥ@���O�ۤv���� = False
      �ĤG���O�ۤv���� = False
      �O�i���ʪ� = False
      ���a�@�ӧQ = False
      �´ѳӧQ = False
      �Ĥ@���� = Empty
      �ĤG���� = Empty
      �Ĥ@���Ѧ�m = Empty
      �ĤG���Ѧ�m = Empty
      �������� = Empty
      �ѥk�W������ = False
      �D��.Show
      Unload �H��
   Else
      Cancel = 1
   End If
End Sub
Private Sub ��_Click(Index As Integer)
   If �w�}�� = False Then '�M�w�C��
      If 16 > ��(Index).Tag And ��(Index).Tag >= 0 Then '�I��´�
         ���a�@.ForeColor = vbBlack
         ���a�G.ForeColor = vbRed
         ���a�@�O�ª� = True
      Else '�I�����
         ���a�@.ForeColor = vbRed
         ���a�G.ForeColor = vbBlack
      End If
      Call ���H
      �w�}�� = True
      ��(Index).Picture = LoadPicture(App.Path & "\" & ��(Index).Tag & ".che")
      �Ѥw½�}(Index) = True
   Else '�w�M�w�C��
      If �Ѥw½�}(Index) = False Then '�Ӵѥ�½�}
         ��(Index).Picture = LoadPicture(App.Path & "\" & ��(Index).Tag & ".che")
         �Ѥw½�}(Index) = True
         Call ���H
      ElseIf �w���Ĥ@���� = True And �Ѥw½�}(Index) = False Then '�ĤG���ѥ�½�}
         If �O�_�t�Y = True Then '�i�t�Y
            GoTo �i�t�Y '���P�w½�}
         Else '���i�t�Y�A���}
            Exit Sub
         End If
         '�٨S��
      ElseIf �w���Ĥ@���� = False Then '�ӴѤw½�}�A�O�Ĥ@��
         �w���Ĥ@���� = True
         �Ĥ@���� = ��(Index).Tag
         �Ĥ@���Ѧ�m = Index
         Call �P�_�Ĥ@���O�_���ۤv����
         If �Ĥ@���O�ۤv���� = False Then '���O�ۤv����
            �w���Ĥ@���� = False '�٨S��
            �Ĥ@���� = Empty
            �Ĥ@���Ѧ�m = Empty
         Else '�O�ۤv����
            �Ĥ@���O�ۤv���� = False
            ���ѹϤ�.Picture = LoadPicture(App.Path & "\" & ��(Index).Tag & ".che")
         End If
      ElseIf �w���Ĥ@���� = True And �Ѥw½�}(Index) = True Then '�ӴѤw½�}�A�O�ĤG��
         If ��(Index).Tag = �Ĥ@���� Then '�M�Ĥ@���ѬۦP�A��������
            �w���Ĥ@���� = False '�٨S��
            �Ĥ@���� = Empty
            �Ĥ@���Ѧ�m = Empty
            ���ѹϤ�.Picture = LoadPicture(App.Path & "\�ť�.che")
         End If
         'If �Ĥ@���� = 9 Or �Ĥ@���� = 10 Then '�ҥ~�G�Ĥ@���ѬO�]
         '   If 32 > ��(Index).Tag And ��(Index).Tag >= 27 Then '�ĤG���ѬO�L
         '       Exit Sub '���P����
         '   End If
         'ElseIf �Ĥ@���� = 25 Or �Ĥ@���� = 26 Then '�ҥ~�G�Ĥ@���ѬO��
         '   If 16 > ��(Index).Tag And ��(Index).Tag >= 15 Then '�ĤG���ѬO��
         '       Exit Sub '���P����
         '   End If
         'End If
�i�t�Y:
         �ĤG���� = ��(Index).Tag
         �ĤG���Ѧ�m = Index
         Call �P�_�ĤG���O�_���ۤv����
         If �ĤG���O�ۤv���� = True Then
            �ĤG���O�ۤv���� = False
            �ĤG���� = Empty
            �ĤG���Ѧ�m = Empty
            Exit Sub '�ĤG���O�ۤv���ѡA����Y�A���}
         Else '�ĤG�����O�ۤv���ѡA�i�H�Y
            Call �ˬd�����Ѫ���m
            If �O�i���ʪ� = False Then
               Exit Sub
            End If
            �O�i���ʪ� = False
            �������� = 0
            Call �ʴѤl
         End If
      End If
   End If
   If Val(����Ѵ�.Caption) = 0 Then '�����C���A�´ѳӧQ
      If ���a�@�O�ª� = True Then '���a�@�ӧQ
         ���a�@�ӧQ = True
         �´ѳӧQ = True
         Call �����C��
      Else '���a�G�ӧQ
         ���a�@�ӧQ = False
         �´ѳӧQ = True
         Call �����C��
      End If
   ElseIf Val(�¤�Ѵ�.Caption) = 0 Then '�����C���A���ѳӧQ
      If ���a�@�O�ª� = False Then '���a�@�ӧQ
         ���a�@�ӧQ = True
         �´ѳӧQ = False
         Call �����C��
      Else '���a�G�ӧQ
         ���a�@�ӧQ = False
         �´ѳӧQ = False
         Call �����C��
      End If
   End If
End Sub
Private Sub �P�_�Ĥ@���O�_���ۤv����()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If 16 > �Ĥ@���� And �Ĥ@���� >= 0 Then '�Ĥ@���O�´�
            �Ĥ@���O�ۤv���� = True
         End If
      Else '���a�@�O����
         If 32 > �Ĥ@���� And �Ĥ@���� >= 16 Then '�Ĥ@���O����
            �Ĥ@���O�ۤv���� = True
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If 32 > �Ĥ@���� And �Ĥ@���� >= 16 Then '�Ĥ@���O����
            �Ĥ@���O�ۤv���� = True
         End If
      Else '���a�G�O�ª�
         If 16 > �Ĥ@���� And �Ĥ@���� >= 0 Then '�Ĥ@���O�´�
            �Ĥ@���O�ۤv���� = True
         End If
      End If
   End If
End Sub
Private Sub �P�_�ĤG���O�_���ۤv����()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If 16 > �ĤG���� And �ĤG���� >= 0 Then '�ĤG���O�´�
            �ĤG���O�ۤv���� = True
         End If
      Else '���a�@�O����
         If 32 > �ĤG���� And �ĤG���� >= 16 Then '�ĤG���O����
            �ĤG���O�ۤv���� = True
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If 32 > �ĤG���� And �ĤG���� >= 16 Then '�ĤG���O����
            �ĤG���O�ۤv���� = True
         End If
      Else '���a�G�O�ª�
         If 16 > �ĤG���� And �ĤG���� >= 0 Then '�ĤG���O�´�
            �ĤG���O�ۤv���� = True
         End If
      End If
   End If
End Sub
Private Sub ���H()
      If �{�b�����a�G = False Then
         �{�b�����a�G = True
         �����a�@��.Visible = False
         �����a�G��.Visible = True
      Else
         �{�b�����a�G = False
         �����a�@��.Visible = True
         �����a�G��.Visible = False
      End If
End Sub
Private Sub �ˬd�����Ѫ���m()
   If �O�_���� = False Then '������
      GoTo ���}���� '�M�@��Ѱʪk�@��
   End If
   If �Ĥ@���� = 9 Or �Ĥ@���� = 10 Or �Ĥ@���� = 25 Or �Ĥ@���� = 26 Then  '�Ĥ@���ѬO�]�ά��A�b������
      If Abs((�Ĥ@���Ѧ�m - �ĤG���Ѧ�m)) < 8 Then '�Ĥ@���ѩM�ĤG���ѦP��
         If ��(�ĤG���Ѧ�m).Tag = -1 Then '�ĤG���ѬO�Ū�
            If �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 1 Or �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 1 Then '�ĤG���ѬO�Ū��B�b�Ĥ@���ѤW��ΤU��
               Call �ʬ����M��
            End If
         Else '�ĤG���Ѥ��O�Ū�
            Temp1 = �Ĥ@���Ѧ�m
            Temp2 = �ĤG���Ѧ�m
            If Temp1 > Temp2 Then
               Temp1 = �ĤG���Ѧ�m
               Temp2 = �Ĥ@���Ѧ�m
            End If
            For i = Temp1 + 1 To Temp2
               If ��(i).Tag <> -1 Then
                  �������� = �������� + 1
               End If
            Next
            If �������� = 2 Then
               Call �ʬ����M��
            End If
         End If
      ElseIf Abs(�Ĥ@���Ѧ�m - �ĤG���Ѧ�m) Mod 8 = 0 Then  '�Ĥ@���ѩM�ĤG���ѦP�C
         If ��(�ĤG���Ѧ�m).Tag = -1 Then '�ĤG���ѬO�Ū�
            If �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 8 Or �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 8 Then '
               Call �ʬ����M��
            End If
         Else '�ĤG���Ѥ��O�Ū�
            Temp1 = �Ĥ@���Ѧ�m
            Temp2 = �ĤG���Ѧ�m
            If Temp1 > Temp2 Then
               Temp1 = �ĤG���Ѧ�m
               Temp2 = �Ĥ@���Ѧ�m
            End If
            For i = Temp1 + 8 To Temp2 Step 8
               If ��(i).Tag <> -1 Then
                  �������� = �������� + 1
               End If
            Next
            If �������� = 2 Then
               Call �ʬ����M��
            End If
         End If
      End If
   Else '�Ĥ@���Ѥ��O�]�ά�
���}����:
      If �Ĥ@���Ѧ�m Mod 8 = �ĤG���Ѧ�m Mod 8 And �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 8 Then '�ĤG���Ѧb�Ĥ@���ѤW��
         �O�i���ʪ� = True
      ElseIf �Ĥ@���Ѧ�m Mod 8 = �ĤG���Ѧ�m Mod 8 And �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 8 Then '�ĤG���Ѧb�Ĥ@���ѤU��
         �O�i���ʪ� = True
      ElseIf �Ĥ@���Ѧ�m = �ĤG���Ѧ�m + 1 Then '�ĤG���Ѧb�Ĥ@���ѥk��
         �O�i���ʪ� = True
      ElseIf �Ĥ@���Ѧ�m = �ĤG���Ѧ�m - 1 Then '�ĤG���Ѧb�Ĥ@���ѥ���
         �O�i���ʪ� = True
      End If
   End If
End Sub
Private Sub �Y���M��()
   �w���Ĥ@���� = False
   �Ĥ@���� = Empty
   �ĤG���� = Empty
   �Ĥ@���Ѧ�m = Empty
   �ĤG���Ѧ�m = Empty
   ���ѹϤ�.Picture = LoadPicture(App.Path & "\�ť�.che")
   Call ���H
End Sub
Private Sub �ʬ����M��()
   If �{�b�����a�G = False Then '�����a�@
      If ���a�@�O�ª� = True Then '���a�@�O�ª�
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         End If
      Else '���a�@�O����
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         End If
      End If
   Else '�����a�G
      If ���a�@�O�ª� = True Then '���a�G�O����
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         End If
      Else '���a�G�O�ª�
         If �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 1 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m + 8 And �Ĥ@���Ѧ�m <> �ĤG���Ѧ�m - 8 Then
            ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         End If
      End If
   End If
   ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".che")
   ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
   ���ѹϤ�.Picture = LoadPicture(App.Path & "\�ť�.che")
   Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
   �Ů�(�Ĥ@���Ѧ�m) = True
   �Ĥ@���� = Empty
   �ĤG���� = Empty
   �Ĥ@���Ѧ�m = Empty
   �ĤG���Ѧ�m = Empty
   �w���Ĥ@���� = False
   �������� = 0
   Call ���H
End Sub
Private Sub �ʴѤl()
   If �{�b�����a�G = False Then
      If ���a�@�O�ª� = True Then
         Call �ʶ´�
      Else '���a�@�O����
         Call �ʬ���
      End If
   Else    '�{�b�����a�G
      If ���a�@�O�ª� = True Then
         Call �ʬ���
      Else '���a�G�O�ª�
         Call �ʶ´�
      End If
   End If
End Sub
Private Sub �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag()
   ��(�ĤG���Ѧ�m).Tag = ��(�Ĥ@���Ѧ�m).Tag
   ��(�Ĥ@���Ѧ�m).Tag = -1
End Sub
Private Sub ���Ѥl�@�ܴѤl�G()
   ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".che")
   ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
End Sub
Private Sub �ʶ´�()
   If �Ĥ@���� = 0 Then '�N
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 27 > �ĤG���� And �ĤG���� >= 16 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 32 > �ĤG���� And �ĤG���� >= 27 Then '�N�Y�L
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 1 Or �Ĥ@���� = 2 Then '�h
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 17 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 16 Then '�h�Y��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 3 Or �Ĥ@���� = 4 Then '�H
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 19 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 16 Or �ĤG���� = 17 Or �ĤG���� = 18 Then '�H�Y�ӥK
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 5 Or �Ĥ@���� = 6 Then '��
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 21 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 21 > �ĤG���� And �ĤG���� >= 16 Then '���Y�ӥK��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 7 Or �Ĥ@���� = 8 Then '��
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 23 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 23 > �ĤG���� And �ĤG���� >= 16 Then '���Y�ӥK����
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 9 Or �Ĥ@���� = 10 Then '�]
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 32 > �ĤG���� And �ĤG���� >= 27 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 27 > �ĤG���� And �ĤG���� >= 16 Then
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf 16 > �Ĥ@���� And �Ĥ@���� >= 11 Then '��
      If 16 > �ĤG���� And �ĤG���� >= 0 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 25 Or �ĤG���� = 26 Then '�Y��
         '����
      ElseIf �ĤG���� = 16 Or �ĤG���� = 27 Or �ĤG���� = 28 Or �ĤG���� = 29 Or �ĤG���� = 30 Or �ĤG���� = 31 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 25 > �ĤG���� And �ĤG���� >= 17 Then    '��Y�K���ϽX
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   End If
End Sub
Private Sub �ʬ���()
   If �Ĥ@���� = 16 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf 11 > �ĤG���� And �ĤG���� >= 0 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf 15 > �ĤG���� And �ĤG���� > 11 Then '�ӦY��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 17 Or �Ĥ@���� = 18 Then '�K
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then   '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Then '�K�Y�N
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 19 Or �Ĥ@���� = 20 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Then '�ۦY�N�h
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 21 Or �Ĥ@���� = 22 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Then '�ۦY�N�h�H
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 23 Or �Ĥ@���� = 24 Then '�X
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then  '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Then '�ۦY�N�h�H��
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf �Ĥ@���� = 25 Or �Ĥ@���� = 26 Then '��
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y����
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 0 Or �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 9 Or �ĤG���� = 10 Then
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   ElseIf 32 > �Ĥ@���� And �Ĥ@���� >= 27 Then '�L
      If 32 > �ĤG���� And �ĤG���� >= 16 Then '�Y�P��
         '�ۤv�ѡA����
      ElseIf �ĤG���� = 9 Or �ĤG���� = 10 Then '�Y�]
         '����
      ElseIf �ĤG���� = 0 Or �ĤG���� = 7 Or �ĤG���� = 8 Or �ĤG���� = 11 Or �ĤG���� = 12 Or �ĤG���� = 13 Or �ĤG���� = 14 Or �ĤG���� = 15 Then '�Y����
         ��(�ĤG���Ѧ�m).Picture = LoadPicture(App.Path & "\" & �Ĥ@���� & ".che")
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         �¤�Ѵ�.Caption = Val(�¤�Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �ĤG���� = 1 Or �ĤG���� = 2 Or �ĤG���� = 3 Or �ĤG���� = 4 Or �ĤG���� = 5 Or �ĤG���� = 6 Then '�L�Y�h�H����
         ��(�Ĥ@���Ѧ�m).Picture = LoadPicture("")
         �Ů�(�Ĥ@���Ѧ�m) = True
         ��(�Ĥ@���Ѧ�m).Tag = -1
         ����Ѵ�.Caption = Val(����Ѵ�.Caption) - 1
         Call �Y���M��
      ElseIf �Ů�(�ĤG���Ѧ�m) = True Then
         Call ���Ѥl�@�ܴѤl�G
         �Ů�(�Ĥ@���Ѧ�m) = True
         �Ů�(�ĤG���Ѧ�m) = False
         Call �ĤG���Ѫ�Tag���Ĥ@���Ѫ�Tag�òM���Ĥ@���Ѫ�Tag
         Call �Y���M��
      End If
   End If
End Sub
Private Sub �����C��()
   �����Ϥ��p�ɾ�.Enabled = True
   ���s��.Visible = True
   If ���a�@�ӧQ = True Then
      �����T��.Caption = Name1P & "�ӧQ"
   Else
      �����T��.Caption = Name2P & "�ӧQ"
   End If
End Sub
Private Sub ���s�}�l_Click()
   Call �~�P
   For i = 0 To 31 '�\�P
      ��(i).Picture = LoadPicture(App.Path & "\�ť�.che")
   Next
   �¤�Ѵ�.Caption = "16"
   ����Ѵ�.Caption = "16"
   �����a�@��.Visible = True
   �����a�G��.Visible = False
   �{�b�����a�G = False
   �w���Ĥ@���� = False
   For i = 0 To 32
      �Ѥw½�}(i) = False
   Next
   For j = 0 To 32
      �Ů�(j) = False
   Next
   �w�}�� = False
   �{�b�����a�G = False
   ���a�@�O�ª� = False
   �Ĥ@���O�ۤv���� = False
   �ĤG���O�ۤv���� = False
   �O�i���ʪ� = False
   ���a�@�ӧQ = False
   �´ѳӧQ = False
   �Ĥ@���� = Empty
   �ĤG���� = Empty
   �Ĥ@���Ѧ�m = Empty
   �ĤG���Ѧ�m = Empty
   �������� = Empty
   ���ѹϤ�.Picture = LoadPicture(App.Path & "\�ť�.che")
End Sub
Private Sub ���v_Click()
   If �{�b�����a�G = False Then '���a�@���v
      ���a�@�ӧQ = False
      Call �����C��
   Else '���a�G���v
      ���a�@�ӧQ = True
      Call �����C��
   End If
End Sub
