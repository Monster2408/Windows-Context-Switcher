'===============================================================================
'[概要]
'  Windows11で追加となったコンテキストメニューを無効化したのを有効化してWindows11仕様に戻す
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

' インストール開始のメッセージボックスを表示する
strMessage = "コンテキストメニューを有効化します。" & vbCr & vbCr

MsgBox strMessage, vbInformation + vbSystemModal, "コンテキストメニュー有効化"

'===============================================================================
' ファイルコピーを実行する
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

If objFileSys.FileExists("新システムへ.bat") Then
 
    ' batファイルを実行する
    objShell.Run "新システムへ.bat",0,True

    ' インストール完了のメッセージボックスを表示する
    strMessage = "コンテキストメニューを有効化しました。" & vbCr & vbCr
    
    MsgBox strMessage, vbInformation + vbSystemModal, "コンテキストメニュー有効化"

Else

    ' EZPartsがインストールされていない、またはWindows 32bit OSの場合はエラーとする
    
    ' 警告メッセージを表示する
    strMessage = "ファイルが正しくコピーされていないため切り替えができません。" & vbCr & vbCr
    
    MsgBox strMessage, vbExclamation + vbSystemModal, "コンテキストメニュー有効化"

End If

'オブジェクトを破棄
Set objShell   = Nothing
Set objFileSys = Nothing

