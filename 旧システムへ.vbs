'===============================================================================
'[概要]
'  Windows11で追加となったコンテキストメニューを無効化してWindows10仕様に戻す
'===============================================================================

Option Explicit

Dim objAppl

'===============================================================================
' 管理者モードで実行するため、VBSを管理者モードで再実行する
'===============================================================================

Set objAppl = Wscript.CreateObject("Shell.Application")
If Wscript.Arguments.Count = 0 then
	objAppl.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
	Wscript.Quit
End If
 

'===============================================================================
' 切替確認のメッセージボックスを表示する
'===============================================================================

' 確認メッセージボックス表示用
Dim strMessage

' 切替開始のメッセージボックスを表示する
strMessage = "コンテキストメニューを無効化します。" & vbCr & vbCr

MsgBox strMessage, vbInformation + vbSystemModal, "コンテキストメニュー無効化"

'===============================================================================
' batを実行する
'===============================================================================

' ファイル処理オブジェクト
Dim objFileSys

' ファイル処理オブジェクトを生成する
Set objFileSys = CreateObject("Scripting.FileSystemObject")

Dim objShell
Set objShell = CreateObject("WScript.Shell")

dim fso, path
set fso = createObject("Scripting.FileSystemObject")
path = fso.GetFile(WScript.ScriptFullName).ParentFolder.Path

objShell.currentDirectory = path & "\bat"

If objFileSys.FileExists("旧システムへ.bat") Then
 
    ' batファイルを実行する
    objShell.Run "旧システムへ.bat",0,True

    ' 切替完了のメッセージボックスを表示する
    strMessage = "コンテキストメニューを無効化しました。" & vbCr & vbCr
    
    MsgBox strMessage, vbInformation + vbSystemModal, "コンテキストメニュー無効化"

Else

    ' BATファイル等が不足していた場合はｴﾗｰ
    
    ' 警告メッセージを表示する
    strMessage = "ファイルが正しくコピーされていないため切り替えができません。" & vbCr & vbCr
    
    MsgBox strMessage, vbExclamation + vbSystemModal, "コンテキストメニュー無効化"

End If

'オブジェクトを破棄
Set objShell   = Nothing
Set objFileSys = Nothing
Set fso = Nothing
