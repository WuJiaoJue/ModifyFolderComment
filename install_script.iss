[Setup]
AppName=ModifyFolderComment PowerShell版
AppVersion=1.1.63p
DefaultDirName={commonpf}\ModifyFolderComment
DefaultGroupName=ModifyFolderComment
OutputBaseFilename=ModifyFolderComment_Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayName=ModifyFolderComment PowerShell版
UninstallDisplayIcon={app}\ModifyFolderComment.ps1

[Files]
Source: "ModifyFolderComment.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "RunModifyFolderComment.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "AFC.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment"; ValueType: string; ValueName: ""; ValueData: "修改文件夹备注"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment\command"; ValueType: string; ValueName: ""; ValueData: "wscript.exe ""{app}\RunModifyFolderComment.vbs"" ""%1"""; Flags: uninsdeletevalue

[Icons]
Name: "{group}\修改文件夹备注 (命令行)"; Filename: "powershell.exe"; Parameters: "-File ""{app}\AFC.ps1"" -h"; WorkingDir: "{app}"
Name: "{group}\卸载 ModifyFolderComment"; Filename: "{uninstallexe}"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
  if MsgBox('此程序需要PowerShell执行权限。安装后可能需要设置执行策略。' + #13#10 + '是否继续安装？', mbConfirmation, MB_YESNO) = IDNO then
    Result := False;
end;

[Run]
Filename: "powershell.exe"; Parameters: "-Command ""Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force"""; Flags: runhidden; Description: "设置PowerShell执行策略"