using System.Reflection;
using System.Runtime.InteropServices;


namespace ModifyFolderComment.Helpers
{
    /// <summary>
    /// 提供文件夹刷新功能，用于修改文件夹备注后触发刷新显示
    /// </summary>
    public static class FolderRefresher
    {
        /// <summary>
        /// 刷新指定文件夹的显示，使其应用新的 desktop.ini 设置
        /// </summary>
        /// <param name="folderPath">要刷新的文件夹路径</param>
        public static void RefreshFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
                return;
            
            var t = new Thread(RefreshFolderStaParameterized);
            t.SetApartmentState(ApartmentState.STA);
            t.Start(folderPath);
            t.Join();
        }
        
        private static void RefreshFolderStaParameterized(object? state)
        {
            if (state is not string folderPath)
                return;
            
            RefreshFolderSta(folderPath);
        }

        private static void RefreshFolderSta(string folderPath)
        {
            var desktopIni = Path.Combine(folderPath, "desktop.ini");
            if (!File.Exists(desktopIni))
            {
                Console.WriteLine("desktop.ini 文件不存在");
                return;
            }

            // 拷贝 desktop.ini 到系统临时目录
            var tempIni = Path.Combine(Path.GetTempPath(), $"desktop_{Guid.NewGuid():N}.ini");
            File.Copy(desktopIni, tempIni, true);

            // 确保临时文件有系统+隐藏属性
            File.SetAttributes(tempIni, FileAttributes.Hidden | FileAttributes.System);

            // 创建 Shell 对象
            var shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null)
            {
                Console.WriteLine("Shell.Application 无法获取");
                return;
            }

            dynamic? shell = Activator.CreateInstance(shellType);
            object? folder = shell?.NameSpace(folderPath);
            if (folder == null)
            {
                Console.WriteLine("无法获取文件夹对象");
                return;
            }

            // 使用 MoveHere 拖入临时 ini 文件，模拟 Explorer 拷贝
            folder.GetType().InvokeMember("MoveHere",
                BindingFlags.InvokeMethod, null, folder, [tempIni, 4 + 16 + 1024]);

            Console.WriteLine("MoveHere 执行完成");
            
            // 等待资源管理器应用 desktop.ini
            Thread.Sleep(100);
            
            // 尝试删除文件夹中的副本（同名或重命名形式）
            try
            {
                var leftovers = Directory.GetFiles(folderPath, "desktop_*.ini");
                foreach (var file in leftovers)
                {
                    File.Delete(file);
                    Console.WriteLine($"已清理残留文件: {Path.GetFileName(file)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理残留文件失败: {ex.Message}");
            }
            
            // 强制通知资源管理器（可选增强）
            SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);
        }

        
        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
    }
}