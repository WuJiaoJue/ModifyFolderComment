param (
    [string]$F,
    [string]$C,
    [switch]$h
)

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
    exit
}

# 检查文件夹是否存在
if (-not (Test-Path $F)) {
    Write-Host "错误：指定的文件夹不存在。"
    exit
}

# 高级桌面ini管理功能 - 与主脚本保持一致
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

# 高级文件夹刷新功能
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
            
            # 尝试删除文件夹中的副本
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

try {
    Set-FolderInfoTip -folderPath $F -comment $C
    Write-Host "已成功为文件夹添加备注。"
}
catch {
    Write-Host "操作失败：$($_.Exception.Message)"
    exit 1
}