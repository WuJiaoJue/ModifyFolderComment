# >>> Invoke-PS2EXE -InputFile "ModifyFolderComment.ps1" -OutputFile "ModifyFolderComment.exe" -NoConsole

param (
    [string]$FolderPath,
    [string]$Comment,
    [switch]$AdminElevated
)

Add-Type -AssemblyName System.Windows.Forms

# 检查是否以管理员身份运行
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# 高级桌面ini管理功能 - 移植自C#版本
function Set-FolderInfoTip {
    param(
        [string]$folderPath,
        [string]$comment
    )
    
    # 获取desktop.ini文件的完整路径
    $iniPath = Join-Path $folderPath 'desktop.ini'
    
    # 读取现有内容或创建新的空列表
    $lines = @()
    if (Test-Path $iniPath) {
        $lines = Get-Content $iniPath -Encoding Unicode
    }
    
    # 查找.ShellClassInfo节
    $sectionIndex = -1
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i].Trim() -eq "[.ShellClassInfo]") {
            $sectionIndex = $i
            break
        }
    }
    
    if ($sectionIndex -eq -1) {
        # 如果节不存在，添加节和InfoTip
        $lines += "[.ShellClassInfo]"
        $lines += "InfoTip=$comment"
    } else {
        # 如果节存在，查找并修改InfoTip，或添加新的
        $replaced = $false
        for ($i = $sectionIndex + 1; $i -lt $lines.Length -and -not $lines[$i].StartsWith("["); $i++) {
            if ($lines[$i].StartsWith("InfoTip=")) {
                $lines[$i] = "InfoTip=$comment"
                $replaced = $true
                break
            }
        }
        if (-not $replaced) {
            # 在节后面插入InfoTip
            $newLines = @()
            $newLines += $lines[0..$sectionIndex]
            $newLines += "InfoTip=$comment"
            if ($sectionIndex + 1 -lt $lines.Length) {
                $newLines += $lines[($sectionIndex + 1)..($lines.Length - 1)]
            }
            $lines = $newLines
        }
    }
    
    # 清除文件只读属性，确保可写入
    if (Test-Path $iniPath) {
        [System.IO.File]::SetAttributes($iniPath, [System.IO.FileAttributes]::Normal)
    }
    
    # 写入更新后的内容
    $lines | Out-File -FilePath $iniPath -Encoding Unicode -Force
    
    # 设置文件为隐藏和系统属性
    [System.IO.File]::SetAttributes($iniPath, [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
    
    # 设置文件夹为系统文件夹
    $dirInfo = Get-Item $folderPath
    $dirInfo.Attributes = $dirInfo.Attributes -bor [System.IO.FileAttributes]::System
    
    # 使用高级刷新功能
    Invoke-FolderRefresh $folderPath
}

# 高级文件夹刷新功能 - 移植自C#版本的FolderRefresher
function Invoke-FolderRefresh {
    param([string]$folderPath)
    
    if (-not (Test-Path $folderPath)) {
        return
    }
    
    $desktopIni = Join-Path $folderPath 'desktop.ini'
    if (-not (Test-Path $desktopIni)) {
        Write-Host "desktop.ini 文件不存在"
        return
    }
    
    try {
        # 拷贝 desktop.ini 到系统临时目录
        $tempIni = Join-Path ([System.IO.Path]::GetTempPath()) "desktop_$([System.Guid]::NewGuid().ToString('N')).ini"
        Copy-Item $desktopIni $tempIni -Force
        
        # 确保临时文件有系统+隐藏属性
        [System.IO.File]::SetAttributes($tempIni, [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
        
        # 创建 Shell 对象并执行 MoveHere 操作
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.NameSpace($folderPath)
        
        if ($folder -ne $null) {
            # 使用 MoveHere 拖入临时 ini 文件，模拟 Explorer 拷贝
            # 参数: 4 (不显示进度) + 16 (响应"是"到所有询问) + 1024 (不显示UI)
            $folder.MoveHere($tempIni, 4 + 16 + 1024)
            Write-Host "MoveHere 执行完成"
            
            # 等待资源管理器应用 desktop.ini
            Start-Sleep -Milliseconds 100
            
            # 尝试删除文件夹中的副本（同名或重命名形式）
            try {
                $leftovers = Get-ChildItem -Path $folderPath -Filter "desktop_*.ini" -Force
                foreach ($file in $leftovers) {
                    Remove-Item $file.FullName -Force
                    Write-Host "已清理残留文件: $($file.Name)"
                }
            }
            catch {
                Write-Host "清理残留文件失败: $($_.Exception.Message)"
            }
        }
        
        # 强制通知资源管理器
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public class RefreshExplorer {
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
"@
        [RefreshExplorer]::SHChangeNotify(0x8000000, 0x1000, [System.IntPtr]::Zero, [System.IntPtr]::Zero)
    }
    catch {
        Write-Host "文件夹刷新失败: $($_.Exception.Message)"
        # 如果高级刷新失败，使用基本的SHChangeNotify
        try {
            [RefreshExplorer]::SHChangeNotify(0x08000000, 0x0000, [System.IntPtr]::Zero, [System.IntPtr]::Zero)
        }
        catch {
            # 忽略刷新失败
        }
    }
}

# 尝试提权并重新调用自身
function Invoke-ElevateAndSetInfoTip {
    param(
        [string]$folderPath,
        [string]$comment
    )
    
    try {
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) {
            $scriptPath = $MyInvocation.MyCommand.Path
        }
        
        $arguments = "-File `"$scriptPath`" -FolderPath `"$folderPath`" -Comment `"$($comment.Replace('"', '\"'))`" -AdminElevated"
        
        Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs -Wait
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("无法请求管理员权限：$($_.Exception.Message)", "权限不足", "OK", "Error")
    }
}

# 获取现有备注信息
function Get-FolderInfoTip {
    param([string]$folderPath)
    
    $iniPath = Join-Path $folderPath 'desktop.ini'
    if (-not (Test-Path $iniPath)) {
        return ""
    }
    
    try {
        $iniContent = Get-Content $iniPath -Encoding Unicode
        $inShellClassInfo = $false
        foreach ($line in $iniContent) {
            if ($line.Trim() -eq "[.ShellClassInfo]") {
                $inShellClassInfo = $true
                continue
            }
            if ($line.StartsWith("[") -and $line.EndsWith("]")) {
                $inShellClassInfo = $false
                continue
            }
            if ($inShellClassInfo -and $line.StartsWith("InfoTip=")) {
                return $line.Substring("InfoTip=".Length)
            }
        }
    }
    catch {
        # 读取失败，返回空字符串
    }
    
    return ""
}

# 检查是否是管理员模式的直接执行
if ($AdminElevated -and $Comment -ne $null) {
    # 直接执行设置操作，不需要UI
    try {
        Set-FolderInfoTip -folderPath $FolderPath -comment $Comment
        Write-Host "管理员模式执行成功"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("以管理员权限修改失败：$($_.Exception.Message)", "错误", "OK", "Error")
    }
    exit
}

# 检查参数
if ([string]::IsNullOrEmpty($FolderPath)) {
    [System.Windows.Forms.MessageBox]::Show("请通过右键菜单或命令行指定文件夹路径。", "参数缺失", "OK", "Warning")
    exit
}

if (-not (Test-Path $FolderPath)) {
    [System.Windows.Forms.MessageBox]::Show("文件夹路径无效。", "错误", "OK", "Error")
    exit
}

# 创建输入框窗口
$form = New-Object System.Windows.Forms.Form
$form.Text = "修改文件夹备注"
$form.Size = New-Object System.Drawing.Size(400,150)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# 添加标签
$label = New-Object System.Windows.Forms.Label
$label.Text = "请输入文件夹的备注信息："
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($label)

# 添加文本框
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(360,20)
$textBox.Location = New-Object System.Drawing.Point(10,50)
$form.Controls.Add($textBox)

# 获取现有备注并显示在文本框中
$existingComment = Get-FolderInfoTip -folderPath $FolderPath
if (-not [string]::IsNullOrWhiteSpace($existingComment)) {
    $textBox.Text = $existingComment
}

# 添加确认按钮
$buttonOk = New-Object System.Windows.Forms.Button
$buttonOk.Text = "确定"
$buttonOk.Location = New-Object System.Drawing.Point(220,80)
$buttonOk.Add_Click({
    $commentText = $textBox.Text.Trim()
    $form.Tag = $commentText
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})
$form.Controls.Add($buttonOk)

# 添加取消按钮
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Text = "取消"
$buttonCancel.Location = New-Object System.Drawing.Point(300,80)
$buttonCancel.Add_Click({
    $form.Tag = $null
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})
$form.Controls.Add($buttonCancel)

# 设置 AcceptButton 和 CancelButton
$form.AcceptButton = $buttonOk
$form.CancelButton = $buttonCancel

# 显示表单并获取结果
$result = $form.ShowDialog()

# 根据DialogResult判断用户操作
if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    exit
}

$comment = $form.Tag

# 即使是空备注也要处理（允许清空备注）
if ($null -eq $comment) {
    exit
}

try {
    Set-FolderInfoTip -folderPath $FolderPath -comment $comment
}
catch [System.UnauthorizedAccessException] {
    # 权限不足，尝试提权
    Invoke-ElevateAndSetInfoTip -folderPath $FolderPath -comment $comment
}
catch [System.IO.IOException] {
    # IO异常，可能是权限问题，尝试提权
    if ($_.Exception.HResult -eq -2147024891) { # E_ACCESSDENIED
        Invoke-ElevateAndSetInfoTip -folderPath $FolderPath -comment $comment
    } else {
        [System.Windows.Forms.MessageBox]::Show("操作失败：$($_.Exception.Message)", "错误", "OK", "Error")
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("操作失败：$($_.Exception.Message)", "错误", "OK", "Error")
}

# 确保脚本不返回任何内容
exit