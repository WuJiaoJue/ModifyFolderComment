Set objShell = CreateObject("Shell.Application")
folderPath = WScript.Arguments(0)
exePath = "{app}\ModifyFolderComment.exe"
objShell.ShellExecute exePath, Chr(34) & folderPath & Chr(34), "", "runas", 1
