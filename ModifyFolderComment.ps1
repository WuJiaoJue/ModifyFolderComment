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
        $iniContent = Get-Content $iniPath -Encoding Unicode
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
    }
    if (-not $iniHash.ContainsKey('.ShellClassInfo')) {
        $iniHash['.ShellClassInfo'] = @{}
    }
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
    # 写回 desktop.ini
    $output = New-Object System.Collections.Generic.List[string]
    foreach ($section in $iniHash.Keys) {
        $output.Add("[$section]")
        foreach ($key in $iniHash[$section].Keys) {
            $value = $iniHash[$section][$key]
            $output.Add("$key=$value")
        }
        $output.Add("")
    }
    Set-Content -Path $iniPath -Value $output -Encoding Unicode -Force
    [System.IO.File]::SetAttributes($iniPath, 'Hidden, System')
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