namespace ModifyFolderComment
{
    /// <summary>
    /// 程序主类，处理程序入口点和初始化逻辑
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// 处理命令行参数，检查管理员权限，并启动主窗体
        /// </summary>
        /// <param name="args">命令行参数，第一个参数应为要修改注释的文件夹路径</param>
        [STAThread]
        private static void Main(string[] args)
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);    // 设置DPI感知，解决高DPI显示器上的字体模糊问题
            Application.EnableVisualStyles();                       // 启用视觉样式
            Application.SetCompatibleTextRenderingDefault(false);   // 设置兼容的文本渲染默认值

            // 获取命令行参数中的文件夹路径
            var folderPath = args.Length > 0 ? args[0] : null;
            // 检查是否提供了文件夹路径
            if (string.IsNullOrEmpty(folderPath))
            {
                // 未提供路径，显示错误信息
                MessageBox.Show("请通过右键菜单或命令行指定文件夹路径。", "参数缺失", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 检查提供的路径是否存在
            if (!Directory.Exists(folderPath))
            {
                // 路径不存在，显示错误信息
                MessageBox.Show("文件夹路径无效。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            // 如果是管理员代执行模式
            if (args.Contains("--admin-elevated"))
            {
                var path = args.FirstOrDefault(arg => !arg.StartsWith("--"));
                var commentIndex = Array.IndexOf(args, "--comment");
                var comment = (commentIndex >= 0 && commentIndex < args.Length - 1) ? args[commentIndex + 1] : "";

                try
                {
                    // 直接调用 InternalSetInfoTip，不需要再次判断权限
                    if (path != null) DesktopIniManager.InternalSetInfoTip(path, comment);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("以管理员权限修改失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }

            // 启动主窗体，传入文件夹路径
            Application.Run(new FolderCommentForm(folderPath));
        }
    }
}