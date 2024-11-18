Set objShell = CreateObject("Shell.Application")
strExe = """E:\tools\MFC\ModifyFolderComment.exe"" """ & WScript.Arguments(0) & """"
objShell.ShellExecute "E:\tools\MFC\ModifyFolderComment.exe", Chr(34) & WScript.Arguments(0) & Chr(34), "", "open", 1