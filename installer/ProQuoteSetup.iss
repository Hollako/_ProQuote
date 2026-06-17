#define MyAppName "ProQuote"
#ifndef MyAppVersion
#define MyAppVersion "0.1.0"
#endif
#define MyAppPublisher "Hollako"
#define MyAppURL "https://github.com/Hollako/_ProQuote"

[Setup]
AppId={{F64E57D7-5664-4FE8-A03E-0212B1E5E06E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\Programs\ProQuote
DefaultGroupName=ProQuote
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=output
OutputBaseFilename=ProQuoteSetup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
SetupLogging=yes
UninstallDisplayIcon={app}\Start ProQuote Installed.bat

[Dirs]
Name: "{localappdata}\ProQuoteData\default"

[Files]
Source: "..\*.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Start ProQuote Installed.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Install Dependencies.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\.streamlit\config.toml"; DestDir: "{app}\.streamlit"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\ProQuote"; Filename: "{app}\Start ProQuote Installed.bat"; WorkingDir: "{app}"
Name: "{group}\ProQuote"; Filename: "{app}\Start ProQuote Installed.bat"; WorkingDir: "{app}"
Name: "{group}\Install or update dependencies"; Filename: "{app}\Install Dependencies.bat"; WorkingDir: "{app}"
Name: "{group}\Uninstall ProQuote"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\Install Dependencies.bat"; Description: "Install/update Python dependencies"; Flags: postinstall runascurrentuser skipifdoesntexist
Filename: "{app}\Start ProQuote Installed.bat"; Description: "Launch ProQuote"; Flags: postinstall nowait unchecked runascurrentuser