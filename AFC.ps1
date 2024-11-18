param (
    [string]$F,
    [string]$C,
    [switch]$h
)

if ($h) {
    Write-Host "使用说明："
    Write-Host "此脚本用于为指定的文件夹添加备注信息。"
    Write-Host "参数："
    Write-Host "  -F `<string>`  要添加备注的文件夹路径。"
    Write-Host "  -C `<string>`  要添加的备注内容。"
    Write-Host "  -h             显示帮助信息。"
    Write-Host "示例："
    Write-Host "  .\AFC.ps1 -F 'C:\示例文件夹' -C '这是备注信息。'"
    exit
}

# 检查是否提供了必要的参数
if (-not $F -or -not $C) {
    Write-Host "错误：缺少必要的参数。请使用 -h 查看使用说明。"
    exit
}

# 检查文件夹是否存在
if (-not (Test-Path $F)) {
    Write-Host "错误：指定的文件夹不存在。"
    exit
}

# 将文件夹属性设置为系统和只读，以便 desktop.ini 生效
$folder = Get-Item $F
$folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::System -bor [System.IO.FileAttributes]::ReadOnly

# 创建或修改 desktop.ini 文件
$iniPath = Join-Path $F 'desktop.ini'

$iniContent = @"
[.ShellClassInfo]
InfoTip=$C
"@

# 以 Unicode 编码写入内容到 desktop.ini
Set-Content -Path $iniPath -Value $iniContent -Encoding Unicode

# 将 desktop.ini 设置为隐藏和系统文件
$iniFile = Get-Item $iniPath
$iniFile.Attributes = $iniFile.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System

Write-Host "已成功为文件夹添加备注。"