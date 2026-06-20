#define MyAppName "ProQuote"
#ifndef MyAppVersion
#define MyAppVersion "0.1.0"
#endif
#define MyAppPublisher "Hollako"
#define MyAppURL "https://github.com/Hollako/_ProQuote"
#ifndef PythonVersion
#define PythonVersion "3.13.14"
#endif
#define PythonInstaller "python-" + PythonVersion + "-amd64.exe"

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
Source: "..\assets\*.png"; DestDir: "{localappdata}\ProQuoteData\default\assets"; Flags: ignoreversion onlyifdoesntexist uninsneveruninstall
Source: "payload\{#PythonInstaller}"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: NeedsPython

[Icons]
Name: "{autodesktop}\ProQuote"; Filename: "{app}\Start ProQuote Installed.bat"; WorkingDir: "{app}"
Name: "{group}\ProQuote"; Filename: "{app}\Start ProQuote Installed.bat"; WorkingDir: "{app}"
Name: "{group}\Install or update dependencies"; Filename: "{app}\Install Dependencies.bat"; WorkingDir: "{app}"
Name: "{group}\Uninstall ProQuote"; Filename: "{uninstallexe}"

[Run]
Filename: "{tmp}\{#PythonInstaller}"; Parameters: "/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1 InstallLauncherAllUsers=0 Include_test=0 Include_doc=0 Shortcuts=0"; StatusMsg: "Installing Python {#PythonVersion}..."; Flags: runhidden waituntilterminated; Check: NeedsPython
Filename: "{cmd}"; Parameters: "/C ""{app}\Install Dependencies.bat"" /silent"; StatusMsg: "Installing ProQuote requirements..."; Flags: runhidden waituntilterminated
Filename: "{app}\Start ProQuote Installed.bat"; Description: "Launch ProQuote"; Flags: postinstall nowait unchecked runascurrentuser

[Code]
function CommandWorks(const Command: String): Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec(
    ExpandConstant('{cmd}'),
    '/C "' + Command + ' >NUL 2>&1"',
    '',
    SW_HIDE,
    ewWaitUntilTerminated,
    ResultCode
  ) and (ResultCode = 0);
end;

function NeedsPython(): Boolean;
begin
  Result := not (
    CommandWorks('py -3 --version') or
    CommandWorks('python --version') or
    FileExists(ExpandConstant('{localappdata}\Programs\Python\Python313\python.exe'))
  );
end;
