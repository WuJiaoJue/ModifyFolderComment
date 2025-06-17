# >>> Invoke-PS2EXE -InputFile "ModifyFolderComment.ps1" -OutputFile "ModifyFolderComment.exe" -NoConsole

param (
    [string]$FolderPath
)

Add-Type -AssemblyName System.Windows.Forms

# 检查是否以管理员身份运行
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    [System.Windows.Forms.MessageBox]::Show('请以管理员身份运行此程序，否则备注可能无法生效。', '权限不足', 'OK', 'Warning')
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

# 设置 AcceptButton 和 CancelButton
$form.AcceptButton = $null  # 暂时设置为 null，稍后指定
$form.CancelButton = $null  # 先设为 null，后面再赋值

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
$existingComment = $null
$iniPath = Join-Path $FolderPath 'desktop.ini'

if (Test-Path $iniPath) {
    try {
        $iniContent = Get-Content $iniPath -Encoding Unicode
        foreach ($line in $iniContent) {
            if ($line -match '^InfoTip\s*=\s*(.*)$') {
                $existingComment = $matches[1].Trim()
                break
            }
        }
    }
    catch {
        # 读取 desktop.ini 失败，忽略
    }
}

if (-not [string]::IsNullOrWhiteSpace($existingComment)) {
    $textBox.Text = $existingComment
}

# 添加确认按钮
$buttonOk = New-Object System.Windows.Forms.Button
$buttonOk.Text = "确定"
$buttonOk.Location = New-Object System.Drawing.Point(220,80)
$buttonOk.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($textBox.Text)) {
        $form.Tag = $textBox.Text
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }
    else {
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    }
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

# 检查是否输入了备注
if ($null -eq $comment) {
    exit
}

try {
    # 设置文件夹属性为系统，以便 desktop.ini 生效
    $folder = Get-Item $FolderPath
    if (-not ($folder.Attributes -band [System.IO.FileAttributes]::System)) {
        $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::System
    }

    # 创建或修改 desktop.ini 文件
    $iniPath = Join-Path $FolderPath 'desktop.ini'
    $iniHash = @{}
    if (Test-Path $iniPath) {
        try {
            # 确保能正确读取 desktop.ini，尝试不同编码
            $iniContent = $null
            try {
                $iniContent = Get-Content $iniPath -Encoding Unicode -ErrorAction Stop
            } catch {
                try {
                    $iniContent = Get-Content $iniPath -Encoding Default -ErrorAction Stop
                } catch {
                    $iniContent = Get-Content $iniPath -ErrorAction Stop
                }
            }

            $currentSection = ''
            foreach ($line in $iniContent) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match '^[\[](.*)[\]]$') {
                    $currentSection = $matches[1]
                    if (-not $iniHash.ContainsKey($currentSection)) {
                        $iniHash[$currentSection] = @{}
                    }
                } elseif ($line -match '^(.*?)=(.*)$') {
                    $key = $matches[1].Trim()
                    $value = $matches[2].Trim()
                    if (-not [string]::IsNullOrEmpty($currentSection)) {
                        $iniHash[$currentSection][$key] = $value
                    }
                }
            }
        } catch {
            # 如果读取失败，创建新的 desktop.ini
            Write-Host "无法读取 desktop.ini，将创建新文件: $_" -ForegroundColor Yellow
        }
    }

    # 确保存在 .ShellClassInfo 节
    if (-not $iniHash.ContainsKey('.ShellClassInfo')) {
        $iniHash['.ShellClassInfo'] = @{}
    }

    # 设置备注信息
    if ([string]::IsNullOrWhiteSpace($comment)) {
        # 备注为空时删除 InfoTip 字段
        $iniHash['.ShellClassInfo'].Remove('InfoTip') | Out-Null
    } else {
        $iniHash['.ShellClassInfo']['InfoTip'] = $comment
    }

    # 清理空节
    $sectionsToRemove = @()
    foreach ($section in $iniHash.Keys) {
        if ($iniHash[$section].Count -eq 0) {
            $sectionsToRemove += $section
        }
    }
    foreach ($section in $sectionsToRemove) {
        $iniHash.Remove($section)
    }

    # 写回 desktop.ini - 确保使用正确的格式和编码
    $output = New-Object System.Collections.Generic.List[string]
    foreach ($section in $iniHash.Keys) {
        $output.Add("[$section]")
        foreach ($key in $iniHash[$section].Keys) {
            $value = $iniHash[$section][$key]
            $output.Add("$key=$value")
        }
        $output.Add("")
    }

    # 确保目录为系统属性
    $folder = Get-Item $FolderPath
    if (-not ($folder.Attributes -band [System.IO.FileAttributes]::System)) {
        $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::System
    }

    # 如果存在旧文件，先删除它
    if (Test-Path $iniPath) {
        Remove-Item -Path $iniPath -Force
    }

    # 写入新的 desktop.ini 文件
    Set-Content -Path $iniPath -Value $output -Encoding Unicode -Force

    # 设置 desktop.ini 属性为隐藏和系统
    [System.IO.File]::SetAttributes($iniPath, [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)

    # 直接在 PowerShell 中实现刷新机制，替代 ForceRefresh.vbs
    try {
        [System.Windows.Forms.MessageBox]::Show('开始执行刷新机制', '调试信息', 'OK', 'Information')

        # 确保文件夹路径不为空且存在
        if (-not [string]::IsNullOrEmpty($FolderPath) -and (Test-Path -Path $FolderPath -PathType Container)) {
            # 设置当前 desktop.ini 文件路径
            $desktopIniPath = Join-Path $FolderPath 'desktop.ini'
            # 使用固定临时文件名，而非随机生成
            $tempIniPath = Join-Path $FolderPath 'desktop.ini.tmp'

            [System.Windows.Forms.MessageBox]::Show("刷新路径：$FolderPath`n设置备注：$comment`ndesktop.ini路径：$desktopIniPath", '调试信息', 'OK', 'Information')

            # 等待文件系统操作完成
            Start-Sleep -Milliseconds 200

            # 获取 Shell.Application COM 对象
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.NameSpace($FolderPath)

            # 使用 Try-Catch 块处理文件操作
            try {
                # 如果 desktop.ini 存在，先复制到临时文件
                if (Test-Path -Path $desktopIniPath -PathType Leaf) {
                    [System.Windows.Forms.MessageBox]::Show("desktop.ini 存在，执行复制操作", '调试信息', 'OK', 'Information')

                    # 先删除可能存在的旧临时文件
                    if (Test-Path -Path $tempIniPath) {
                        Remove-Item -Path $tempIniPath -Force -ErrorAction SilentlyContinue
                    }

                    # 复制文件（不立即设置其属性）
                    Copy-Item -Path $desktopIniPath -Destination $tempIniPath -Force -ErrorAction Stop

                    # 确认临时文件创建成功后再操作
                    Start-Sleep -Milliseconds 100
                    if (Test-Path -Path $tempIniPath -PathType Leaf) {
                        # 删除原始 desktop.ini (先删除目标后移动，避免权限问题)
                        Remove-Item -Path $desktopIniPath -Force -ErrorAction SilentlyContinue

                        # 检查删除是否成功
                        if (-not (Test-Path -Path $desktopIniPath)) {
                            [System.Windows.Forms.MessageBox]::Show("原始 desktop.ini 已删除，准备重命名临时文件", '调试信息', 'OK', 'Information')
                        }

                        # 先通过重命名移动文件，避免使用MoveHere导致的复杂性
                        Rename-Item -Path $tempIniPath -NewName "desktop.ini" -Force

                        # 确保设置了正确的文件属性
                        if (Test-Path -Path $desktopIniPath -PathType Leaf) {
                            [System.IO.File]::SetAttributes($desktopIniPath, [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
                            [System.Windows.Forms.MessageBox]::Show("重命名成功并设置了属性", '调试信息', 'OK', 'Information')

                            # 通过Shell的方式刷新缓存
                            $folder.MoveHere($desktopIniPath, 4+16+1024)
                            [System.Windows.Forms.MessageBox]::Show("已执行 MoveHere 刷新操作", '调试信息', 'OK', 'Information')
                        }
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("desktop.ini 不存在，无需刷新", '调试信息', 'OK', 'Information')
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("文件操作失败: $_", '调试错误', 'OK', 'Error')

                # 紧急恢复：确保保留desktop.ini文件
                if ((-not (Test-Path -Path $desktopIniPath)) -and (Test-Path -Path $tempIniPath)) {
                    try {
                        Rename-Item -Path $tempIniPath -NewName "desktop.ini" -Force
                        if (Test-Path -Path $desktopIniPath) {
                            [System.IO.File]::SetAttributes($desktopIniPath, [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
                            [System.Windows.Forms.MessageBox]::Show("紧急恢复完成", '调试信息', 'OK', 'Information')
                        }
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("紧急恢复失败: $_", '调试错误', 'OK', 'Error')
                    }
                }
            }
            finally {
                # 清理临时文件（如果还存在）
                if (Test-Path -Path $tempIniPath) {
                    try {
                        Remove-Item -Path $tempIniPath -Force -ErrorAction SilentlyContinue
                    }
                    catch {
                        # 忽略清理错误
                    }
                }

                # 释放 COM 对象
                try {
                    if ($null -ne $folder) {
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folder) | Out-Null
                    }
                    if ($null -ne $shell) {
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
                    }
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
                catch {
                    # 忽略 COM 对象释放错误
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("文件夹路径无效: $FolderPath", '调试错误', 'OK', 'Error')
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("刷新 shell 缓存失败: $_", '调试错误', 'OK', 'Error')
        # 错误不会阻止主要功能
    }

    # 强制资源管理器刷新
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class RefreshExplorer {
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
"@
    [RefreshExplorer]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)
}
catch {
    [System.Windows.Forms.MessageBox]::Show('操作失败：' + $_.Exception.Message, '错误', 'OK', 'Error')
    exit
}

# 确保脚本不返回任何内容
exit