#define AppName "DocGen"
#define AppVersion "1.0.0"
#define AppPublisher "DocGen"
#define AppExeName "DocGen.exe"
#define IncludeVCRedist 0

[Setup]
AppId={{A9E8B19A-4631-4E48-A54F-6C7C0D7C6E2B}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=dist_installer
OutputBaseFilename=DocGenSetup_{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#AppExeName}

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务"; Flags: unchecked

[Files]
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "word_template\*"; DestDir: "{app}\word_template"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "安装包使用说明.txt"; DestDir: "{app}"; Flags: ignoreversion

#if IncludeVCRedist
Source: "redist\vcredist_x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "redist\vcredist_x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
#endif

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\卸载 {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\DocGen"

#if IncludeVCRedist
[Run]
Filename: "{tmp}\vcredist_x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "正在安装 Visual C++ 运行时..."; Flags: waituntilterminated; Check: Is64BitInstallMode and ShouldInstallVCRedist64
Filename: "{tmp}\vcredist_x86.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "正在安装 Visual C++ 运行时..."; Flags: waituntilterminated; Check: ShouldInstallVCRedist32
#endif

[Code]
function IsVCRedistInstalled(Arch: string): Boolean;
var
  Installed: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\' + Arch, 'Installed', Installed) then
    Result := (Installed = 1)
  else if RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\' + Arch, 'Installed', Installed) then
    Result := (Installed = 1);
end;

function ShouldInstallVCRedist32(): Boolean;
begin
  Result := not IsVCRedistInstalled('x86');
end;

function ShouldInstallVCRedist64(): Boolean;
begin
  Result := not IsVCRedistInstalled('x64');
end;
