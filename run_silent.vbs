' Folder2Print - Silent Runner
' This VBScript runs the Python script without showing a console window
' Use this for autostart scenarios

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.Run "pythonw.exe folder2print.py", 0, False
