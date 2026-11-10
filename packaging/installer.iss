; Builds "Manuscript Check Setup.exe". Installs for the signed-in user only, so no administrator rights are needed.
#define AppName "Manuscript Check"
#ifndef AppVersion
  #define AppVersion "1.0"
#endif

[Setup]
AppId={{8F3C2B1A-5D4E-4C7B-9A1E-2F6D3C8B7A90}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Gideon Gulmatico
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=Manuscript Check Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\Manuscript Check.exe
UninstallDisplayName={#AppName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes

[Messages]
WelcomeLabel2=This installs Manuscript Check for checking theses against the school's format.

[Tasks]
Name: "desktopicon"; Description: "Put a Manuscript Check icon on the desktop"

[Files]
Source: "..\dist\Manuscript Check\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion createallsubdirs
Source: "MicrosoftEdgeWebview2Setup.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: NeedsWebView2

[Icons]
Name: "{userprograms}\{#AppName}"; Filename: "{app}\Manuscript Check.exe"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\Manuscript Check.exe"; Tasks: desktopicon

[Run]
Filename: "{tmp}\MicrosoftEdgeWebview2Setup.exe"; Parameters: "/silent /install"; StatusMsg: "Installing Microsoft Edge WebView2, which the program window needs"; Check: NeedsWebView2; Flags: waituntilterminated
Filename: "{app}\Manuscript Check.exe"; Description: "Open Manuscript Check now"; Flags: nowait postinstall skipifsilent

[Code]
function HasWebView2(Root: Integer; Key: String): Boolean;
var
  Version: String;
begin
  Result := RegQueryStringValue(Root, Key, 'pv', Version) and (Version <> '') and (Version <> '0.0.0.0');
end;

function NeedsWebView2: Boolean;
begin
  Result := not (HasWebView2(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}')
              or HasWebView2(HKLM, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}')
              or HasWebView2(HKCU, 'Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'));
end;
