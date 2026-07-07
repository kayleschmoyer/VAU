[Setup]
AppName=VAST Auto Updater
AppVersion=1.0.0
AppPublisher=VAST Retail
AppSupportURL=https://github.com/kayleschmoyer/VAU
AppUpdatesURL=https://github.com/kayleschmoyer/VAU/releases
DefaultDirName={autopf}\VAST Auto Updater
DefaultGroupName=VAST Auto Updater
OutputDir=Output
OutputBaseFilename=VastAutoUpdater-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=..\VastAutoUpdater\vastretail_icon.ico
UninstallDisplayIcon={app}\VastAutoUpdater.exe
PrivilegesRequired=admin
WizardStyle=modern

[Files]
Source: "..\VastAutoUpdater\bin\Release\VastAutoUpdater.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\VastAutoUpdater\bin\Release\VastAutoUpdater.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\VastAutoUpdater\bin\Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\VAST Auto Updater"; Filename: "{app}\VastAutoUpdater.exe"
Name: "{autodesktop}\VAST Auto Updater"; Filename: "{app}\VastAutoUpdater.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"
Name: "scheduledtask"; Description: "Run silent update check daily at 2:00 AM"; GroupDescription: "Automated updates:"; Flags: checkedonce

[Run]
; Create the scheduled task if the user opted in
Filename: "schtasks.exe"; Parameters: "/Create /TN ""VAST Auto Updater"" /TR """"""{app}\VastAutoUpdater.exe"" silent"" /SC DAILY /ST 02:00 /RL HIGHEST /F"; Flags: runhidden; Tasks: scheduledtask
; Offer to launch after install
Filename: "{app}\VastAutoUpdater.exe"; Description: "Launch VAST Auto Updater"; Flags: nowait postinstall skipifsilent shellexec

[UninstallRun]
; Remove the scheduled task on uninstall
Filename: "schtasks.exe"; Parameters: "/Delete /TN ""VAST Auto Updater"" /F"; Flags: runhidden
