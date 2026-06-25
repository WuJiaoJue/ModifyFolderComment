# >>> Invoke-PS2EXE -InputFile "ModifyFolderComment.ps1" -OutputFile "ModifyFolderComment.exe" -NoConsole -requireAdmin

param (
    [string]$FolderPath,
    [string]$Comment,
    [switch]$AdminElevated
)

# 导入核心模块
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $scriptDir 'FolderCommentCore.psm1'
if (-not (Test-Path $modulePath)) {
    # 尝试从当前目录加载
    $modulePath = Join-Path (Get-Location) 'FolderCommentCore.psm1'
}
Import-Module $modulePath -Force

# 设置日志（可选，可通过 Set-LogFile 自定义路径）
$logPath = Join-Path $env:TEMP "ModifyFolderComment_$(Get-Date -Format 'yyyyMMdd').log"
Set-LogFile -Path $logPath

Add-Type -AssemblyName System.Windows.Forms

# 检查是否以管理员身份运行
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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

# 检查是否是管理员模式的直接执行
if ($AdminElevated -and $null -ne $Comment) {
    try {
        Set-FolderInfoTip -FolderPath $FolderPath -Comment $Comment
        Write-OperationLog -Operation "SetComment" -FolderPath $FolderPath -Comment $Comment -Success $true
        Write-Host "管理员模式执行成功"
    }
    catch {
        Write-OperationLog -Operation "SetComment" -FolderPath $FolderPath -Comment $Comment -Success $false -ErrorMessage $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("以管理员权限修改失败：$($_.Exception.Message)", "错误", "OK", "Error")
    }
    exit
}

# 检查参数
if ([string]::IsNullOrEmpty($FolderPath)) {
    [System.Windows.Forms.MessageBox]::Show("请通过右键菜单或命令行指定文件夹路径。", "参数缺失", "OK", "Warning")
    exit
}

# 标准化路径（处理末尾反斜杠、解析相对路径等）
$FolderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($FolderPath)

# 验证路径是否真实存在
if (-not (Test-Path $FolderPath)) {
    [System.Windows.Forms.MessageBox]::Show("文件夹路径无效。", "错误", "OK", "Error")
    exit
}

# 验证是否为文件夹
$item = Get-Item $FolderPath -Force
if ($item -isnot [System.IO.DirectoryInfo]) {
    [System.Windows.Forms.MessageBox]::Show("指定的路径不是文件夹。", "错误", "OK", "Error")
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
$existingComment = Get-FolderInfoTip -FolderPath $FolderPath
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
    Set-FolderInfoTip -FolderPath $FolderPath -Comment $comment
    Write-OperationLog -Operation "SetComment" -FolderPath $FolderPath -Comment $comment -Success $true
}
catch [System.UnauthorizedAccessException] {
    Write-OperationLog -Operation "SetComment" -FolderPath $FolderPath -Comment $comment -Success $false -ErrorMessage "UnauthorizedAccessException"
    Invoke-ElevateAndSetInfoTip -folderPath $FolderPath -comment $comment
}
catch [System.IO.IOException] {
    if ($_.Exception.HResult -eq -2147024891) {
        Write-OperationLog -Operation "SetComment" -FolderPath $FolderPath -Comment $comment -Success $false -ErrorMessage "AccessDenied"
        Invoke-ElevateAndSetInfoTip -folderPath $FolderPath -comment $comment
    }
    else {
        Write-OperationLog -Operation "SetComment" -FolderPath $FolderPath -Comment $comment -Success $false -ErrorMessage $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("操作失败：$($_.Exception.Message)", "错误", "OK", "Error")
    }
}
catch {
    Write-OperationLog -Operation "SetComment" -FolderPath $FolderPath -Comment $comment -Success $false -ErrorMessage $_.Exception.Message
    [System.Windows.Forms.MessageBox]::Show("操作失败：$($_.Exception.Message)", "错误", "OK", "Error")
}

exit
