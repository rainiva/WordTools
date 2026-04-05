; Word工具箱 - InnoSetup 安装脚本
; 编码: UTF-8 with BOM

#define MyAppName "Word工具箱"
#define MyAppVersion "1.2.0.0"
#define MyAppPublisher "WordTools"
#define MyAppProgId "WordTools.ThisAddIn"
#define MyAppDLL "WordTools.dll"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\WordToolbox
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=Output
OutputBaseFilename=WordToolbox_Setup
SetupIconFile=
Compression=lzma
SolidCompression=yes
WizardStyle=modern
LanguageDetectionMethod=uilanguage
ShowLanguageDialog=yes

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "WordTools\bin\Debug\{#MyAppDLL}"; DestDir: "{app}"; Flags: ignoreversion
Source: "WordTools\bin\Debug\Extensibility.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "WordTools\bin\Debug\stdole.dll"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Word Addins 注册表项
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Word工具箱"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Word工具箱插件"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"

[Run]
; 安装后注册 COM 组件
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/codebase ""{app}\{#MyAppDLL}"""; StatusMsg: "正在注册 COM 组件..."; Flags: runhidden

[UninstallRun]
; 卸载前反注册 COM 组件
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/unregister ""{app}\{#MyAppDLL}"""; StatusMsg: "正在反注册 COM 组件..."; Flags: runhidden

[Code]
function InitializeSetup(): Boolean;
begin
  Result := true;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    MsgBox('安装完成！' + #13#10 + #13#10 + 
           '请完全关闭 Microsoft Word（包括后台进程），然后重新打开 Word 以启用插件。' + #13#10 + #13#10 +
           '启用后，Word 顶部会出现"Word工具箱"选项卡。', 
           mbInformation, MB_OK);
  end;
end;

function InitializeUninstall(): Boolean;
begin
  Result := true;
  MsgBox('卸载前请确保 Microsoft Word 已完全关闭。', mbInformation, MB_OK);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    MsgBox('卸载完成！' + #13#10 + #13#10 + 
           '插件已从系统中移除。', 
           mbInformation, MB_OK);
  end;
end;
