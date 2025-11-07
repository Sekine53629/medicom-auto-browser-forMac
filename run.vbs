Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' スクリプトのディレクトリを取得
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPythonScript = strScriptPath & "\main.py"

' PowerShellで実行（UTF-8エンコーディング付き、ウィンドウを閉じない）
strCommand = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -Command ""& {[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; [Console]::InputEncoding = [System.Text.Encoding]::UTF8; python '" & strPythonScript & "';}"""

' ウィンドウを表示して実行
objShell.Run strCommand, 1, False
