Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 获取文件夹路径参数
If WScript.Arguments.Count = 0 Then
    MsgBox "请通过右键菜单调用此程序。", vbExclamation, "参数缺失"
    WScript.Quit
End If

folderPath = WScript.Arguments(0)

' 检查文件夹是否存在
If Not objFSO.FolderExists(folderPath) Then
    MsgBox "文件夹路径无效。", vbCritical, "错误"
    WScript.Quit
End If

' 获取当前脚本目录
scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
psScriptPath = objFSO.BuildPath(scriptDir, "ModifyFolderComment.ps1")

' 检查PowerShell脚本是否存在
If Not objFSO.FileExists(psScriptPath) Then
    MsgBox "找不到 ModifyFolderComment.ps1 文件。", vbCritical, "错误"
    WScript.Quit
End If

' 构建PowerShell命令
psCommand = "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & """ -FolderPath """ & folderPath & """"

' 执行PowerShell脚本
objShell.ShellExecute "powershell.exe", "-ExecutionPolicy Bypass -File """ & psScriptPath & """ -FolderPath """ & folderPath & """", "", "open", 1
