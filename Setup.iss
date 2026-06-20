; Word工具箱 - Inno Setup 安装脚本
; 当前版本仅正式支持 64 位 Microsoft Word。

; MyAppVersion is synced from version.json via sync-version.ps1.
#ifndef MyAppVersion
  #define MyAppVersion "1.3.0"
#endif
#define MyAppName "Word工具箱"
#define MyAppPublisher "WordTools"
#define MyAppProgId "WordTools.ThisAddIn"
#define MyAppDLL "WordTools.dll"

#if !Defined(SourceConfiguration)
  #define SourceConfiguration "Release"
#endif

#if Defined(ARCH_X86) && Defined(ARCH_X64)
  #error 不能同时定义 ARCH_X86 和 ARCH_X64
#endif

#if !Defined(ARCH_X86) && !Defined(ARCH_X64)
  #define ARCH_X64
#endif

#if Defined(ARCH_X86)
  #define BuildLabel "32 位环境（当前版本暂不支持）"
  #define OutputBaseName "WordToolbox_Setup_" + MyAppVersion + "_x86"
  #define DefaultDirNameValue "{autopf32}\WordToolbox"
  #define ArchitecturesAllowedValue "x86compatible"
#else
  #define BuildLabel "64 位 Microsoft Word"
  #define OutputBaseName "WordToolbox_Setup_" + MyAppVersion + "_x64"
  #define DefaultDirNameValue "{autopf64}\WordToolbox"
  #define ArchitecturesAllowedValue "x64compatible"
  #define RegAsmDirectory "{win}\Microsoft.NET\Framework64\v4.0.30319"
#endif

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={#DefaultDirNameValue}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
UsedUserAreasWarning=no
ArchitecturesAllowed={#ArchitecturesAllowedValue}
#if Defined(ARCH_X64)
ArchitecturesInstallIn64BitMode=x64compatible
#endif
OutputDir=dist
OutputBaseFilename={#OutputBaseName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
LanguageDetectionMethod=uilanguage
ShowLanguageDialog=yes
UninstallDisplayIcon={app}\{#MyAppDLL}

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
#if Defined(ARCH_X64)
Source: "WordTools\bin\{#SourceConfiguration}\{#MyAppDLL}"; DestDir: "{app}"; Flags: ignoreversion
Source: "WordTools\bin\{#SourceConfiguration}\Extensibility.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "WordTools\bin\{#SourceConfiguration}\stdole.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "INSTALLATION.md"; DestDir: "{app}"; DestName: "安装说明.txt"; Flags: ignoreversion
#endif

[Registry]
#if Defined(ARCH_X64)
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Word工具箱"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Word工具箱插件"
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#MyAppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"
#endif

[Run]
#if Defined(ARCH_X64)
Filename: "{#RegAsmDirectory}\RegAsm.exe"; Parameters: "/codebase ""{app}\{#MyAppDLL}"""; StatusMsg: "正在注册 COM 组件..."; Flags: runhidden
#endif

[UninstallRun]
#if Defined(ARCH_X64)
Filename: "{#RegAsmDirectory}\RegAsm.exe"; Parameters: "/unregister ""{app}\{#MyAppDLL}"""; StatusMsg: "正在反注册 COM 组件..."; Flags: runhidden; RunOnceId: "WordTools.Unregister"
#endif

[Code]
function InitializeSetup(): Boolean;
begin
#if Defined(ARCH_X86)
  MsgBox(
    '当前版本仅支持 64 位 Microsoft Word。' + #13#10 + #13#10 +
    '暂不支持 32 位 Word、32 位 WPS、64 位 WPS。' + #13#10 +
    '请勿继续安装当前 x86 安装包。',
    mbCriticalError,
    MB_OK);
  Result := False;
#else
  MsgBox(
    '当前安装包仅支持 64 位 Microsoft Word。' + #13#10 + #13#10 +
    '暂不支持 32 位 Word、32 位 WPS、64 位 WPS。' + #13#10 +
    '如果你的宿主不是 64 位 Microsoft Word，请取消安装。',
    mbInformation,
    MB_OK);
  Result := True;
#endif
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
#if Defined(ARCH_X64)
  if CurStep = ssPostInstall then
  begin
    MsgBox(
      '安装完成！' + #13#10 + #13#10 +
      '当前版本仅支持 64 位 Microsoft Word。' + #13#10 +
      '暂不支持 32 位 Word、32 位 WPS、64 位 WPS。' + #13#10 + #13#10 +
      '请完全关闭 Word（包括后台进程）后重新打开。',
      mbInformation,
      MB_OK);
  end;
#endif
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  MsgBox('卸载前请确保 Word 已完全关闭。', mbInformation, MB_OK);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
#if Defined(ARCH_X64)
  if CurUninstallStep = usPostUninstall then
  begin
    MsgBox(
      '卸载完成！' + #13#10 + #13#10 +
      '插件已移除。如 Word 仍在运行，请重新启动 Word。',
      mbInformation,
      MB_OK);
  end;
#endif
end;
