[Setup]
AppName=ModifyFolderComment
AppVersion=1.1.1
DefaultDirName={commonpf}\ModifyFolderComment
DefaultGroupName=ModifyFolderComment
OutputBaseFilename=install
Compression=lzma2
SolidCompression=yes

[InstallDelete]
Type: files; Name: "{app}\RunModifyFolderComment.vbs.bak"

[Files]
; 包含可执行文件和脚本
Source: "ModifyFolderComment.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "RunModifyFolderComment.vbs"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\ModifyFolderComment"; Filename: "{app}\ModifyFolderComment.exe"

[Registry]
; 导入注册表文件
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment"; ValueType: string; ValueName: ""; ValueData: "修改文件夹备注"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Directory\shell\ModifyFolderComment\command"; ValueType: string; ValueName: ""; ValueData: "wscript.exe ""{app}\RunModifyFolderComment.vbs"" ""%1"""; Flags: uninsdeletevalue

[Code]
procedure ReplaceInFile(const FileName, SearchString, ReplaceString: string);
var
  Lines: TArrayOfString;
  i: Integer;
begin
  if LoadStringsFromFile(FileName, Lines) then
  begin
    for i := 0 to GetArrayLength(Lines) - 1 do
    begin
      StringChangeEx(Lines[i], SearchString, ReplaceString, True);
    end;
    SaveStringsToFile(FileName, Lines, False);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    ReplaceInFile(ExpandConstant('{app}\RunModifyFolderComment.vbs'), '{app}', ExpandConstant('{app}'));
  end;
end;