' Khởi động SAP - Outlook - Teams Bridge Worker chạy ngầm 100% không hiện cửa sổ CMD đen
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = scriptDir

cmdPath = scriptDir & "\START_BRIDGE.cmd"
WshShell.Run "cmd.exe /c """ & cmdPath & """", 0, False
