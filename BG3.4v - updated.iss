#define MyAppName "Bank Guarantee Radar System"
#define MyAppVersion "3.4"
#define MyAppPublisher "RAJAT DUBEY"
; ⚠️ COMMON GENERIC EXE NAME FOR ALL FUTURE VERSIONS
#define MyAppExeName "Bank Guarantee Radar System.exe"
#define MyAppID "{D25B7680-6A24-460A-B92A-A73505E04894}"

[Setup]
AppId={{#MyAppID}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

; --- FILE PROPERTIES METADATA ---
VersionInfoVersion={#MyAppVersion}.0.0
VersionInfoProductVersion={#MyAppVersion}.0.0
VersionInfoCompany=NORTH WESTERN RAILWAY - BIKANER DIVISION
VersionInfoDescription=Bank Guarantee Radar System Setup
VersionInfoCopyright=Copyright (C) 2026 Shri Rajat Dubey
VersionInfoProductName=Bank Guarantee Radar System

DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

DisableProgramGroupPage=yes

; --- AUTO-UPDATE OVERWRITE SETTINGS ---
DirExistsWarning=no
CloseApplications=yes
CloseApplicationsFilter=*.exe
RestartApplications=no
PrivilegesRequired=admin

LicenseFile=C:\Users\Rajat Dubey\Desktop\BGP\License_Agreement.txt
InfoBeforeFile=C:\Users\Rajat Dubey\Desktop\BGP\Info_Before_Install.txt
InfoAfterFile=C:\Users\Rajat Dubey\Desktop\BGP\Info_After_Install.txt

OutputDir=C:\Users\Rajat Dubey\Desktop\BGP
OutputBaseFilename=BG_Radar_System_v{#MyAppVersion}_Setup
SetupIconFile=C:\Users\Rajat Dubey\Desktop\BGP\app_icon.ico

SolidCompression=yes
WizardStyle=modern dynamic windows11

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
SetupAppTitle=Setup - {#MyAppName} v{#MyAppVersion}
SetupWindowTitle=Setup - {#MyAppName} v{#MyAppVersion}
WizardInstalling=Updating
InstallingLabel=Please wait while Setup updates {#MyAppName} on your computer.

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

; 🧹 DYNAMIC CLEANUP: Wildcard removes ANY older version-tagged EXEs dynamically
[InstallDelete]
Type: files; Name: "{app}\Bank Guarantee Radar System (v*).exe"
Type: files; Name: "{app}\Bank Guarantee Radar System v*.exe"

[Files]
Source: "C:\Users\Rajat Dubey\Desktop\BGP\dist\Bank Guarantee Radar System\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app_icon.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app_icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
var
  IsUpdateMode: Boolean;

// 🛑 DYNAMIC PROCESS KILLER: Terminates any app process matching the name pattern
procedure KillAppProcesses();
var
  ResultCode: Integer;
begin
  // Kill main fixed executable
  Exec('taskkill.exe', '/f /im "' + ExpandConstant('{#MyAppExeName}') + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  
  // Dynamic filter query to kill any running EXE containing "Bank Guarantee Radar"
  Exec('taskkill.exe', '/f /fi "IMAGENAME eq Bank Guarantee Radar*"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  
  Sleep(500); // 0.5s pause to ensure Windows releases all file lock handles
end;

procedure ParseVersion(VerStr: String; var Major, Minor, Build, Rev: Integer);
var
  P: Integer;
begin
  Major := 0; Minor := 0; Build := 0; Rev := 0;
  
  P := Pos('.', VerStr);
  if P > 0 then begin
    Major := StrToIntDef(Copy(VerStr, 1, P - 1), 0);
    VerStr := Copy(VerStr, P + 1, Length(VerStr));
  end else begin
    Major := StrToIntDef(VerStr, 0);
    Exit;
  end;

  P := Pos('.', VerStr);
  if P > 0 then begin
    Minor := StrToIntDef(Copy(VerStr, 1, P - 1), 0);
    VerStr := Copy(VerStr, P + 1, Length(VerStr));
  end else begin
    Minor := StrToIntDef(VerStr, 0);
    Exit;
  end;

  P := Pos('.', VerStr);
  if P > 0 then begin
    Build := StrToIntDef(Copy(VerStr, 1, P - 1), 0);
    Rev := StrToIntDef(Copy(VerStr, P + 1, Length(VerStr)), 0);
  end else begin
    Build := StrToIntDef(VerStr, 0);
  end;
end;

function CompareVersions(Ver1, Ver2: String): Integer;
var
  Maj1, Min1, Bld1, Rev1: Integer;
  Maj2, Min2, Bld2, Rev2: Integer;
begin
  ParseVersion(Ver1, Maj1, Min1, Bld1, Rev1);
  ParseVersion(Ver2, Maj2, Min2, Bld2, Rev2);

  if Maj1 <> Maj2 then begin if Maj1 > Maj2 then Result := 1 else Result := -1; Exit; end;
  if Min1 <> Min2 then begin if Min1 > Min2 then Result := 1 else Result := -1; Exit; end;
  if Bld1 <> Bld2 then begin if Bld1 > Bld2 then Result := 1 else Result := -1; Exit; end;
  if Rev1 <> Rev2 then begin if Rev1 > Rev2 then Result := 1 else Result := -1; Exit; end;
  
  Result := 0;
end;

function GetInstalledVersion(out InstalledVer, UninstallString: String): Boolean;
var
  RegKey: String;
begin
  Result := False;
  RegKey := 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppID}_is1';

  if RegQueryStringValue(HKLM64, RegKey, 'DisplayVersion', InstalledVer) then
  begin
    RegQueryStringValue(HKLM64, RegKey, 'UninstallString', UninstallString);
    Result := True;
  end
  else if RegQueryStringValue(HKCU64, RegKey, 'DisplayVersion', InstalledVer) then
  begin
    RegQueryStringValue(HKCU64, RegKey, 'UninstallString', UninstallString);
    Result := True;
  end;
end;

function InitializeSetup(): Boolean;
var
  InstalledVer, UninstallStr: String;
  CmpRes: Integer;
begin
  Result := True;
  IsUpdateMode := False;

  // Kill running app dynamically before file overwrite operations
  KillAppProcesses();

  if WizardSilent() then
  begin
    Exit;
  end;

  if GetInstalledVersion(InstalledVer, UninstallStr) then
  begin
    CmpRes := CompareVersions(InstalledVer, ExpandConstant('{#MyAppVersion}'));

    if CmpRes < 0 then
    begin
      IsUpdateMode := True;
      if MsgBox('A previous version (' + InstalledVer + ') is installed.' + #13#10 +
                'Do you want to UPDATE to Version ' + ExpandConstant('{#MyAppVersion}') + '?',
                mbConfirmation, MB_YESNO) = IDNO then
      begin
        Result := False;
        Exit;
      end;
    end;
  end;
end;

procedure InitializeWizard();
begin
  if IsUpdateMode then
  begin
    WizardForm.Caption := 'Updating - ' + ExpandConstant('{#MyAppName}') + ' v' + ExpandConstant('{#MyAppVersion}');
    WizardForm.NextButton.Caption := 'Update';
  end;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if IsUpdateMode then
  begin
    if (PageID = wpLicense) or (PageID = wpInfoBefore) or (PageID = wpInfoAfter) or (PageID = wpSelectTasks) then
    begin
      Result := True;
    end;
  end;
end;