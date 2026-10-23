' Khởi động SAP - Outlook - Teams Bridge Worker chạy ngầm 100% không hiện cửa sổ CMD đen
Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.Run "python main.py", 0, False
