'===============================================================================
'[�T�v]
'  Windows11�Œǉ��ƂȂ����R���e�L�X�g���j���[�𖳌��������̂�L��������Windows11�d�l�ɖ߂�
'===============================================================================

Option Explicit

Dim objAppl

'===============================================================================
' �Ǘ��҃��[�h�Ŏ��s���邽�߁AVBS���Ǘ��҃��[�h�ōĎ��s����
'===============================================================================

Set objAppl = Wscript.CreateObject("Shell.Application")
If Wscript.Arguments.Count = 0 then
	objAppl.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
	Wscript.Quit
End If
 

'===============================================================================
' �ؑ֊m�F�̃��b�Z�[�W�{�b�N�X��\������
'===============================================================================

' �m�F���b�Z�[�W�{�b�N�X�\���p
Dim strMessage

' �C���X�g�[���J�n�̃��b�Z�[�W�{�b�N�X��\������
strMessage = "�R���e�L�X�g���j���[��L�������܂��B" & vbCr & vbCr

MsgBox strMessage, vbInformation + vbSystemModal, "�R���e�L�X�g���j���[�L����"

'===============================================================================
' �t�@�C���R�s�[�����s����
'===============================================================================

' �t�@�C�������I�u�W�F�N�g
Dim objFileSys

' �t�@�C�������I�u�W�F�N�g�𐶐�����
Set objFileSys = CreateObject("Scripting.FileSystemObject")

Dim objShell
Set objShell = CreateObject("WScript.Shell")

dim fso, path
set fso = createObject("Scripting.FileSystemObject")
path = fso.GetFile(WScript.ScriptFullName).ParentFolder.Path

objShell.currentDirectory = path & "\bat"

If objFileSys.FileExists("�V�V�X�e����.bat") Then
 
    ' bat�t�@�C�������s����
    objShell.Run "�V�V�X�e����.bat",0,True

    ' �C���X�g�[�������̃��b�Z�[�W�{�b�N�X��\������
    strMessage = "�R���e�L�X�g���j���[��L�������܂����B" & vbCr & vbCr
    
    MsgBox strMessage, vbInformation + vbSystemModal, "�R���e�L�X�g���j���[�L����"

Else

    ' EZParts���C���X�g�[������Ă��Ȃ��A�܂���Windows 32bit OS�̏ꍇ�̓G���[�Ƃ���
    
    ' �x�����b�Z�[�W��\������
    strMessage = "�t�@�C�����������R�s�[����Ă��Ȃ����ߐ؂�ւ����ł��܂���B" & vbCr & vbCr
    
    MsgBox strMessage, vbExclamation + vbSystemModal, "�R���e�L�X�g���j���[�L����"

End If

'�I�u�W�F�N�g��j��
Set objShell   = Nothing
Set objFileSys = Nothing

