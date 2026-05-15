# FolderCommentCore.psm1
# 核心模块：文件夹备注的读写和刷新功能

using namespace System.Runtime.InteropServices

#region Native Methods
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class FolderCommentNative {
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
"@
#endregion

#region Comment Escaping
function Escape-InfoTipComment {
    param([string]$Comment)
    # 转义注释内容，防止破坏 ini 格式
    # 处理换行符（替换为空格，因为 InfoTip 通常是单行显示）
    $escaped = $Comment -replace "[\r\n]+", " "
    # 如果用户输入包含 InfoTip= 前缀，转义等号防止被识别为键
    $escaped = $escaped -replace "^(InfoTip=)", "InfoTip\="
    return $escaped
}
#endregion

#region Core Functions
function Get-FolderInfoTip {
    <#
    .SYNOPSIS
        获取文件夹的备注信息
    .PARAMETER FolderPath
        文件夹路径
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    $iniPath = Join-Path $FolderPath 'desktop.ini'
    if (-not (Test-Path $iniPath)) {
        return ""
    }

    try {
        $iniContent = Get-Content $iniPath -Encoding Unicode -ErrorAction Stop
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

function Set-FolderInfoTip {
    <#
    .SYNOPSIS
        设置文件夹的备注信息
    .PARAMETER FolderPath
        文件夹路径
    .PARAMETER Comment
        备注内容
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath,
        [Parameter(Mandatory = $true)]
        [string]$Comment
    )

    $iniPath = Join-Path $FolderPath 'desktop.ini'

    # 转义注释内容，防止注入
    $escapedComment = Escape-InfoTipComment -Comment $Comment

    # 读取现有内容或创建新的空列表
    $lines = @()
    if (Test-Path $iniPath) {
        $lines = @($(Get-Content $iniPath -Encoding Unicode -ErrorAction Stop))
    }

    # 查找 [.ShellClassInfo] 节
    $sectionIndex = -1
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i].Trim() -eq "[.ShellClassInfo]") {
            $sectionIndex = $i
            break
        }
    }

    if ($sectionIndex -eq -1) {
        # 节不存在，添加
        $lines += "[.ShellClassInfo]"
        $lines += "InfoTip=$escapedComment"
    }
    else {
        # 节存在，查找或插入 InfoTip
        $replaced = $false
        for ($i = $sectionIndex + 1; $i -lt $lines.Length -and -not $lines[$i].StartsWith("["); $i++) {
            if ($lines[$i].StartsWith("InfoTip=")) {
                $lines[$i] = "InfoTip=$escapedComment"
                $replaced = $true
                break
            }
        }
        if (-not $replaced) {
            $newLines = @()
            $newLines += $lines[0..$sectionIndex]
            $newLines += "InfoTip=$escapedComment"
            if ($sectionIndex + 1 -lt $lines.Length) {
                $newLines += $lines[($sectionIndex + 1)..($lines.Length - 1)]
            }
            $lines = $newLines
        }
    }

    # 清除只读属性
    if (Test-Path $iniPath) {
        [System.IO.File]::SetAttributes($iniPath, [System.IO.FileAttributes]::Normal)
    }

    # 写入更新后的内容（Unicode LE with BOM）
    $bytes = [System.Text.Encoding]::Unicode.GetPreamble() +
              ([System.Text.Encoding]::Unicode.GetBytes(($lines -join "`n") + "`n"))
    [System.IO.File]::WriteAllBytes($iniPath, $bytes)

    # 设置隐藏和系统属性
    [System.IO.File]::SetAttributes($iniPath, [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)

    # 设置文件夹为系统文件夹
    $dirInfo = Get-Item $FolderPath
    $dirInfo.Attributes = $dirInfo.Attributes -bor [System.IO.FileAttributes]::System

    # 执行刷新
    Invoke-FolderRefresh -FolderPath $FolderPath
}
#endregion

#region Folder Refresh with STA and Retry
function Invoke-FolderRefresh {
    <#
    .SYNOPSIS
        刷新文件夹显示，使 desktop.ini 更改立即生效
    .PARAMETER FolderPath
        文件夹路径
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    if (-not (Test-Path $FolderPath)) {
        return
    }

    $desktopIni = Join-Path $FolderPath 'desktop.ini'
    if (-not (Test-Path $desktopIni)) {
        Write-Verbose "desktop.ini 文件不存在"
        return
    }

    # 使用后台作业在 STA 线程执行刷新
    $job = Start-Job -ScriptBlock {
        param($folderPath, $desktopIniPath)

        Add-Type -AssemblyName System.Windows.Forms

        $shell = $null
        try {
            # 拷贝 desktop.ini 到系统临时目录
            $tempIni = Join-Path ([System.IO.Path]::GetTempPath()) "desktop_$([System.Guid]::NewGuid().ToString('N')).ini"
            Copy-Item $desktopIniPath $tempIni -Force

            # 确保临时文件有系统+隐藏属性
            [System.IO.File]::SetAttributes($tempIni, [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)

            # 创建 Shell 对象
            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace($folderPath)

            if ($folder -ne $null) {
                # 使用 MoveHere 拖入临时 ini 文件，模拟 Explorer 拷贝
                # 参数: 4 (不显示进度) + 16 (响应"是"到所有询问) + 1024 (不显示UI)
                $folder.MoveHere($tempIni, 4 + 16 + 1024)

                # 等待资源管理器应用 desktop.ini
                Start-Sleep -Milliseconds 150
            }
        }
        finally {
            # 释放 COM 对象
            if ($shell -ne $null) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
            }
        }

        # 强制通知资源管理器
        [FolderCommentNative]::SHChangeNotify(0x8000000, 0x1000, [IntPtr]::Zero, [IntPtr]::Zero)
    } -ArgumentList $FolderPath, $desktopIni

    # 等待作业完成（超时 5 秒）
    $completed = Wait-Job -Job $job -Timeout 5
    if ($completed) {
        $result = Receive-Job -Job $job
        if ($result) { Write-Verbose $result }
    }
    else {
        Write-Verbose "刷新操作超时"
    }
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

    # 清理残留文件
    try {
        Get-ChildItem -Path $FolderPath -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "desktop*.ini" -or $_.Name -like "ini*.tmp" -or $_.Name -like "~$*" } |
            ForEach-Object {
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                Write-Verbose "已清理残留文件: $($_.Name)"
            }
    }
    catch {
        Write-Verbose "清理残留文件失败: $($_.Exception.Message)"
    }
}
#endregion

#region Logging
$script:LogFile = $null

function Set-LogFile {
    param([string]$Path)
    $script:LogFile = $Path
}

function Write-OperationLog {
    param(
        [string]$Operation,
        [string]$FolderPath,
        [string]$Comment,
        [bool]$Success,
        [string]$ErrorMessage = $null
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp | $Operation | Path: $FolderPath | Success: $Success"
    if ($ErrorMessage) {
        $logEntry += " | Error: $ErrorMessage"
    }

    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }
    Write-Verbose $logEntry
}
#endregion

Export-ModuleMember -Function @(
    'Get-FolderInfoTip',
    'Set-FolderInfoTip',
    'Invoke-FolderRefresh',
    'Escape-InfoTipComment',
    'Set-LogFile',
    'Write-OperationLog'
)
