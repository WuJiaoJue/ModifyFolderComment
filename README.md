# ModifyFolderComment 

---
### windouws右键文件夹添加备注信息

![Snipaste_2024-11-19_00-16-04.png](images/Snipaste_2024-11-19_00-16-04.png)

![Snipaste_2024-11-19_00-17-03.png](images/Snipaste_2024-11-19_00-17-03.png)

![Snipaste_2024-11-19_00-18-02.png](images/Snipaste_2024-11-19_00-18-02.png)

## 安装

1. 下载最新版本的安装程序（`install.exe`）。
2. 双击运行安装程序，按照提示完成安装。

## 文件

安装程序包含以下文件：

- `ModifyFolderComment.exe`: 主程序文件。
- `RunModifyFolderComment.vbs`: 用于运行主程序的脚本文件。
- `AddContextMenuOption.reg`: 注册表文件，用于添加右键菜单选项。

_`ModifyFolderComment.exe`为`ModifyFolderComment.ps1`打包而成的可执行文件，不信任的自行下载`ModifyFolderComment.ps1`文件打包。_
> 打包命令
> ```powershell
> Invoke-PS2EXE -InputFile "ModifyFolderComment.ps1" -OutputFile "ModifyFolderComment.exe" -NoConsole -requireAdmin
> ```


## 运行

在资源管理器中，右键单击文件夹，选择 "修改文件夹备注" 选项。在弹出的对话框中输入备注信息，然后单击 "确定"。

## 卸载

要卸载 ModifyFolderComment 应用程序，请执行以下步骤：

1. 打开控制面板，选择 "程序和功能"。
2. 找到 "ModifyFolderComment" 项目，右键单击并选择 "卸载"。

卸载程序将自动删除所有安装的文件和注册表项。

## 打包与安装

- 用 Inno Setup 打开 install_script.iss，点击“编译”生成 install.exe。
- install.exe 会自动包含主程序和脚本，并注册右键菜单。
- 右键菜单注册表也可用 AddContextMenuOption.reg 手动导入。

> install_script.iss 关键片段：
> ```ini
> [Files]
> Source: "ModifyFolderComment.exe"; DestDir: "{app}"
> Source: "RunModifyFolderComment.vbs"; DestDir: "{app}"
> [Registry]
> Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment"; ValueData: "修改文件夹备注"
> Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment\command"; ValueData: "wscript.exe \"{app}\RunModifyFolderComment.vbs\" \"%1\""
> ```
