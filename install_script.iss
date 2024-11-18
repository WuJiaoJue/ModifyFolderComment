; Inno Setup 脚本示例
[Setup]
AppName=ModifyFolderComment
AppVersion=1.0
DefaultDirName={commonpf}\ModifyFolderComment
DefaultGroupName=ModifyFolderComment
OutputBaseFilename=install
Compression=lzma2
SolidCompression=yes

[Files]
; 包含可执行文件和脚本
Source: "E:\tools\MFC\ModifyFolderComment.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "E:\tools\MFC\RunModifyFolderComment.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "AddContextMenuOption.reg"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\ModifyFolderComment"; Filename: "{app}\ModifyFolderComment.exe"

[Registry]
; 导入注册表文件
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment"; ValueType: string; ValueName: ""; ValueData: "修改文件夹备注"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment\command"; ValueType: string; ValueName: ""; ValueData: "wscript.exe ""{app}\RunModifyFolderComment.vbs"" ""%1"""; Flags: uninsdeletevalue

[Run]
; 可选：运行注册表导入
Filename: "regedit.exe"; Parameters: "/s ""{app}\AddContextMenuOption.reg"""; Flags: runhidden