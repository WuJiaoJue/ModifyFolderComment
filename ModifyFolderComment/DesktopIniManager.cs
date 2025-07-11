using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ModifyFolderComment.Helpers;


namespace ModifyFolderComment;

/// <summary>
/// 处理 desktop.ini 文件的读写操作，管理文件夹注释(InfoTip)
/// </summary>
public static class DesktopIniManager
{
    public static string GetInfoTip(string folderPath)
    {
        var iniPath = Path.Combine(folderPath, "desktop.ini");
        if (!File.Exists(iniPath)) return "";

        var lines = File.ReadAllLines(iniPath, Encoding.Unicode);
        string? currentSection = null;
        foreach (var line in lines)
        {
            if (line.StartsWith($"[") && line.EndsWith($"]"))
            {
                currentSection = line.Trim('[', ']');
            }
            else if (currentSection == ".ShellClassInfo" && line.StartsWith("InfoTip="))
            {
                return line["InfoTip=".Length..];
            }
        }
        return "";
    }
    
    public static void SetInfoTip(string folderPath, string comment)
    {
        try
        {
            InternalSetInfoTip(folderPath, comment);
        }
        catch (UnauthorizedAccessException)
        {
            TryElevateAndSetInfoTip(folderPath, comment);
        }
        catch (IOException ex) when (IsAccessDenied(ex))
        {
            TryElevateAndSetInfoTip(folderPath, comment);
        }
    }

    /// <summary>
    /// 设置指定文件夹的注释文本(InfoTip)
    /// </summary>
    /// <param name="folderPath">文件夹路径</param>
    /// <param name="comment">注释文本</param>
    public static void InternalSetInfoTip(string folderPath, string comment)
    {
        // 获取desktop.ini文件的完整路径
        var iniPath = Path.Combine(folderPath, "desktop.ini");
        
        // 读取现有内容或创建新的空列表
        var lines = File.Exists(iniPath)
            ? new List<string>(File.ReadAllLines(iniPath, Encoding.Unicode))
            : new List<string>();

        // 查找.ShellClassInfo节
        var sectionIndex = lines.FindIndex(l => l.Trim() == "[.ShellClassInfo]");
        if (sectionIndex == -1)
        {
            // 如果节不存在，添加节和InfoTip
            lines.Add("[.ShellClassInfo]");
            lines.Add("InfoTip=" + comment);
        }
        else
        {
            // 如果节存在，查找并修改InfoTip，或添加新的
            var replaced = false;
            for (var i = sectionIndex + 1; i < lines.Count && !lines[i].StartsWith("["); i++)
            {
                if (lines[i].StartsWith("InfoTip="))
                {
                    lines[i] = "InfoTip=" + comment;
                    replaced = true;
                    break;
                }
            }
            if (!replaced)
                lines.Insert(sectionIndex + 1, "InfoTip=" + comment);
        }
        
        // 清除文件只读属性，确保可写入
        if (File.Exists(iniPath))
        {
            File.SetAttributes(iniPath, FileAttributes.Normal);
        }

        // 写入更新后的内容
        using (var fs = new FileStream(iniPath, FileMode.Create, FileAccess.Write))
        using (var sw = new StreamWriter(fs, new UnicodeEncoding(false, true))) // Little-endian + BOM
        {
            foreach (var line in lines)
                sw.WriteLine(line);
        }
        // 设置文件为隐藏和系统属性，符合Windows对此文件的标准处理
        File.SetAttributes(iniPath, FileAttributes.Hidden | FileAttributes.System);

        // 设置文件夹为系统文件夹
        var dirInfo = new DirectoryInfo(folderPath);
        dirInfo.Attributes |= FileAttributes.System;
        
        // 通知系统刷新文件夹属性，使更改立即生效
        // SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
        FolderRefresher.RefreshFolder(folderPath);
        SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero); // SHCNE_ASSOCCHANGED
    }
    
    /// <summary>
    /// 检查是否是由于访问被拒绝引起的 IOException
    /// </summary>
    private static bool IsAccessDenied(IOException ex)
    {
        return ex.HResult == -2147024891; // 0x80070005: E_ACCESSDENIED
    }
    
    /// <summary>
    /// 提权并重新调用自身
    /// </summary>
    private static void TryElevateAndSetInfoTip(string folderPath, string comment)
    {
        try
        {
            var exePath = Environment.ProcessPath;
            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true,
                Verb = "runas", // 请求管理员权限
                Arguments = $"\"{folderPath}\" --admin-elevated --comment \"{comment.Replace("\"", "\\\"")}\""
            };
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            MessageBox.Show("无法请求管理员权限：" + ex.Message, "权限不足", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// 从Windows Shell32.dll导入的函数，用于通知系统刷新资源管理器
    /// </summary>
    [DllImport("shell32.dll", SetLastError = true)]
    private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}