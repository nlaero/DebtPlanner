; When this installer is compiled, refresh the first-run test snapshot if APP_VERSION changed.
#define FirstRunSync Exec("python", "sync_first_run_test.py", SourcePath, 1, 0)
#if FirstRunSync != 0
  #error "Could not update first_run_test_debt_payoff.py. Install Python and run the compile from this folder."
#endif

[Setup]
AppId={{7C4E9A21-5B18-4F6D-A3E0-91D2C8B64F17}
SetupIconFile=debt_payoff.ico
AppName=Debt Payoff Planner
AppVersion=2.22
VersionInfoVersion=2.22.0.0
AppPublisher=Debt Payoff Planner
DefaultDirName={autopf}\Debt Payoff Planner
DefaultGroupName=Debt Payoff Planner
DisableProgramGroupPage=yes
OutputDir=installer\Version 2.x
OutputBaseFilename=DebtPayoffPlannerSetup-2.22
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
CloseApplicationsFilter=Debt Payoff Planner.exe
RestartApplications=no
UninstallDisplayIcon={app}\Debt Payoff Planner.exe
SetupLogging=no

[Files]
Source: "debt_payoff.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist-2.22\Debt Payoff Planner\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "refresh_icons.ps1"; Flags: dontcopy

[Icons]
Name: "{group}\Debt Payoff Planner"; Filename: "{app}\Debt Payoff Planner.exe"; WorkingDir: "{app}"; IconFilename: "{app}\debt_payoff.ico"; IconIndex: 0
Name: "{autodesktop}\Debt Payoff Planner"; Filename: "{app}\Debt Payoff Planner.exe"; WorkingDir: "{app}"; IconFilename: "{app}\debt_payoff.ico"; IconIndex: 0

[Run]
Filename: "{app}\Debt Payoff Planner.exe"; Description: "Launch Debt Payoff Planner"; Flags: nowait postinstall

[Code]
function ExePath: String;
begin
  Result := ExpandConstant('{app}\Debt Payoff Planner.exe');
end;

procedure RemoveOldShortcut(const ShortcutPath: String);
begin
  if FileExists(ShortcutPath) then
    DeleteFile(ShortcutPath);
end;

procedure RefreshInstalledIcons;
var
  ResultCode: Integer;
  Ps, Params: String;
begin
  ExtractTemporaryFile('refresh_icons.ps1');
  { The setup program is 32-bit. Explorer is 64-bit, so the redraw has to
    come from the 64-bit PowerShell, or the desktop keeps the old icon. }
  Ps := ExpandConstant('{sysnative}\WindowsPowerShell\v1.0\powershell.exe');
  if not FileExists(Ps) then
    Ps := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
  Params := '-NoProfile -ExecutionPolicy Bypass -File "' + ExpandConstant('{tmp}\refresh_icons.ps1') +
    '" -IconPath "' + ExpandConstant('{app}\debt_payoff.ico') +
    '" -ExePath "' + ExePath + '"';
  Exec(Ps, Params, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    RemoveOldShortcut(ExpandConstant('{autodesktop}\Debt Payoff Planner.lnk'));
    RemoveOldShortcut(ExpandConstant('{userdesktop}\Debt Payoff Planner.lnk'));
    RemoveOldShortcut(ExpandConstant('{commondesktop}\Debt Payoff Planner.lnk'));
    RemoveOldShortcut(ExpandConstant('{group}\Debt Payoff Planner.lnk'));
    RemoveOldShortcut(ExpandConstant('{userprograms}\Debt Payoff Planner\Debt Payoff Planner.lnk'));
    RemoveOldShortcut(ExpandConstant('{commonprograms}\Debt Payoff Planner\Debt Payoff Planner.lnk'));
  end;
  if CurStep = ssPostInstall then
    RefreshInstalledIcons;
end;
