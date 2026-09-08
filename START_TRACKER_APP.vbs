Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = folder

ps1 = folder & "\START_TRACKER_APP.ps1"

' Run the PowerShell launcher hidden.
shell.Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1 & """", 0, False
