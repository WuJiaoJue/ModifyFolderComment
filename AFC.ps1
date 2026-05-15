param (
    [string]$F,
    [string]$C,
    [switch]$h
)

# 导入核心模块
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $scriptDir 'FolderCommentCore.psm1'
if (-not (Test-Path $modulePath)) {
    # 尝试从当前目录加载
    $modulePath = Join-Path (Get-Location) 'FolderCommentCore.psm1'
}
Import-Module $modulePath -Force

# 设置日志
$logPath = Join-Path $env:TEMP "ModifyFolderComment_$(Get-Date -Format 'yyyyMMdd').log"
Set-LogFile -Path $logPath

# 显示帮助信息
if ($h) {
    Write-Host "使用说明："
    Write-Host "此脚本用于为指定的文件夹添加备注信息。"
    Write-Host "参数："
    Write-Host "  -F <string>  要添加备注的文件夹路径。"
    Write-Host "  -C <string>  要添加的备注内容。"
    Write-Host "  -h           显示帮助信息。"
    Write-Host "示例："
    Write-Host "  .\AFC.ps1 -F 'C:\示例文件夹' -C '这是备注信息。'"
    exit
}

# 检查是否提供了必要的参数
if (-not $F -or -not $C) {
    Write-Host "错误：缺少必要的参数。请使用 -h 查看使用说明。"
    exit 1
}

# 标准化路径
$F = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($F)

# 检查文件夹是否存在
if (-not (Test-Path $F)) {
    Write-Host "错误：指定的文件夹不存在。"
    exit 1
}

# 验证是否为文件夹
$item = Get-Item $F -Force
if ($item -isnot [System.IO.DirectoryInfo]) {
    Write-Host "错误：指定的路径不是文件夹。"
    exit 1
}

try {
    Set-FolderInfoTip -FolderPath $F -Comment $C
    Write-OperationLog -Operation "SetComment" -FolderPath $F -Comment $C -Success $true
    Write-Host "已成功为文件夹添加备注。"
}
catch {
    Write-OperationLog -Operation "SetComment" -FolderPath $F -Comment $C -Success $false -ErrorMessage $_.Exception.Message
    Write-Host "操作失败：$($_.Exception.Message)"
    exit 1
}
