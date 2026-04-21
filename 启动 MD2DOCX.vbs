Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

strDir = fso.GetParentFolderName(WScript.ScriptFullName)

WshShell.Run "cmd /c ""cd /d " & strDir & " && uv sync --quiet 2>nul && start /b uv run main.py""", 0, False
