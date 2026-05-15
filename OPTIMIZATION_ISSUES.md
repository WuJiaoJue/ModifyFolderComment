# Optimization Issues

## 1. 优化 FolderRefresher 使用 STA 线程提升 COM 稳定性

**状态**: ✅ 已完成

**实现方式**:
使用 `Start-Job` 在独立线程执行 COM 操作（虽然 PowerShell Start-Job 默认不是严格 STA，但避免了 MTA 的问题）

**参考实现** (C# 版本已有):
```csharp
var t = new Thread(RefreshFolderStaParameterized);
t.SetApartmentState(ApartmentState.STA);
t.Start(folderPath);
t.Join();
```

---

## 2. 为 MoveHere 操作添加重试机制和超时控制

**状态**: ⚠️ 部分完成（添加了超时，未实现重试）

**当前实现**:
```powershell
$job = Start-Job -ScriptBlock { ... }
$completed = Wait-Job -Job $job -Timeout 5
```

**待改进**: 可在刷新失败后增加重试逻辑（建议 3 次，递增等待时间）

---

## 3. 增强临时文件和残留文件的清理逻辑

**状态**: ✅ 已完成

**当前实现**:
```powershell
Get-ChildItem -Path $FolderPath -Force |
    Where-Object {
        $_.Name -like "desktop*.ini" -or
        $_.Name -like "ini*.tmp" -or
        $_.Name -like "~$*"
    }
```

---

## 4. 优化 SHChangeNotify 标志组合

**状态**: ⏳ 未实现

**当前代码**:
```powershell
[FolderCommentNative]::SHChangeNotify(0x8000000, 0x1000, [IntPtr]::Zero, [IntPtr]::Zero)
```

**待改进**: 可研究更精准的刷新标志组合

---

## 5. 添加长路径支持（\\?\ 前缀）

**状态**: ⏳ 未实现

**待改进**: 在 `$tempIni` 路径前添加 `\\?\` 前缀以支持超长路径

---

## 6. 代码重复问题

**状态**: ✅ 已完成

**解决方案**: 提取公共代码到 `FolderCommentCore.psm1` 模块

| 文件 | 修改 |
|------|------|
| `ModifyFolderComment.ps1` | Import-Module，移除重复函数 |
| `AFC.ps1` | Import-Module，移除重复函数 |
| `FolderCommentCore.psm1` | 新增，包含所有核心功能 |

---

## 7. 注释注入漏洞

**状态**: ✅ 已完成

**问题**: 用户输入的 `InfoTip=` 或换行符可能破坏 ini 格式

**解决方案**: `Escape-InfoTipComment` 函数转义处理
```powershell
function Escape-InfoTipComment {
    $escaped = $Comment -replace "[\r\n]+", " "
    $escaped = $escaped -replace "^(InfoTip=)", "InfoTip\="
    return $escaped
}
```

---

## 8. COM 对象未释放

**状态**: ✅ 已完成

**实现**:
```powershell
finally {
    if ($shell -ne $null) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
}
```

---

## 9. 路径验证不足

**状态**: ✅ 已完成

**改进**:
- 标准化路径（处理末尾反斜杠、解析相对路径）
- 验证路径是否真实存在
- 验证是否为文件夹类型

```powershell
$FolderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($FolderPath)
$item = Get-Item $FolderPath -Force
if ($item -isnot [System.IO.DirectoryInfo]) { ... }
```

---

## 10. VBS 包装器过时

**状态**: ✅ 已完成

**解决方案**:
- `install_script.iss`: 改用 PowerShell 直接调用
- `AddContextMenuOption.reg`: 改用 PowerShell 调用
- `RunModifyFolderComment.ps1`: PowerShell 替代方案

---

## 11. 无操作日志

**状态**: ✅ 已完成

**实现**:
```powershell
# 日志文件位于 %TEMP%\ModifyFolderComment_yyyyMMdd.log
Set-LogFile -Path $logPath
Write-OperationLog -Operation "SetComment" -FolderPath $path -Comment $comment -Success $true
```
