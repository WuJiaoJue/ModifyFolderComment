Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' 检查是否有参数
If WScript.Arguments.Count = 0 Then
    WScript.Quit
End If

' 获取文件夹路径参数
' strFolderPath = WScript.Arguments(0)
strFolderPath = "W:\workspace\pyspace\ModifyFolderComment\Output"

' 检查文件夹路径是否存在
If strFolderPath <> "" And objFSO.FolderExists(strFolderPath) Then
    ' 设置当前 desktop.ini 文件路径
    strDesktopIniPath = objFSO.BuildPath(strFolderPath, "desktop.ini")
    ' 设置临时目录路径
    strTempFolderPath = objFSO.GetSpecialFolder(2).Path
    ' 设置临时 desktop.ini 路径
    strTempIniPath = objFSO.BuildPath(strTempFolderPath, "desktop.ini")

    ' 创建一个Shell.Application对象实例，这个对象提供了与Windows外壳交互的方法和属性
    set shell = CreateObject("Shell.Application")
    ' 使用Shell.Application对象的NameSpace方法来获取指定文件夹路径的命名空间对象
    ' 这个命名空间对象代表了文件夹，可以用来访问文件夹的属性、图标、文件等
    set folder = shell.NameSpace(strFolderPath)

    ' 检查 desktop.ini 文件是否存在
    If objFSO.FileExists(strDesktopIniPath) Then
        objFSO.MoveFile strDesktopIniPath, strTempIniPath
        folder.MoveHere strTempIniPath, 4+16+1024
        ' objFSO.MoveFile strTempIniPath, strDesktopIniPath '这个移动无用，不会触发刷新
    Else
        ' 创建新的 desktop.ini 文件内容
        strIniContent = "[.ShellClassInfo]" & vbCrLf &_
                        "IconResource=C:\Windows\System32\SHELL32.dll,4" & vbCrLf

        ' 创建新的 desktop.ini 文件
        Set objTempFile = objFSO.CreateTextFile(strTempIniPath, True)
        objTempFile.Write strIniContent
        objTempFile.Close

        ' 设置 desktop.ini 文件的隐藏和系统属性
        Set objFile = objFSO.GetFile(strTempIniPath)
        objFile.Attributes = objFile.Attributes + 3 ' Hidden + System
        ' 移动到当前文件夹路径
        folder.MoveHere strTempIniPath, 4+16+1024

        ' 删除临时 desktop.ini 文件___手动删除可以触发刷新，这个删除不行。。。
        objFSO.DeleteFile strDesktopIniPath, True
    End If
End If