Set objShell = CreateObject("Shell.Application")
strExe = """{app}\ModifyFolderComment.exe"" """ & WScript.Arguments(0) & """"
objShell.ShellExecute "{app}\ModifyFolderComment.exe", Chr(34) & WScript.Arguments(0) & Chr(34), "", "open", 1