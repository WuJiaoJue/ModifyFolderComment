[Setup]
AppName=ModifyFolderComment
AppVersion=1.1.56c
DefaultDirName={commonpf}\ModifyFolderComment
DefaultGroupName=ModifyFolderComment
OutputBaseFilename=install
Compression=lzma2
SolidCompression=yes

[Files]
Source: "bin\Release\net8.0-windows\win-x64\publish\ModifyFolderComment.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\ModifyFolderComment"; Filename: "{app}\ModifyFolderComment.exe"

[Registry]
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment"; ValueType: string; ValueName: ""; ValueData: "修改文件夹备注"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment\command"; ValueType: string; ValueName: ""; ValueData: """{app}\ModifyFolderComment.exe"" ""%1"""; Flags: uninsdeletevalue