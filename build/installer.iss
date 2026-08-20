; ============================================================================
;  Inno Setup script  --  Procalc Hydraulics
;
;  Packages the PyInstaller one-dir output (dist\Procalc\*) into a Windows
;  installer.  Build order:
;     1)  pyinstaller procalc.spec --clean --noconfirm   ->  dist\Procalc\
;     2)  iscc installer.iss                              ->  Output\ProcalcSetup-<ver>.exe
;  (build.ps1 runs both steps for you.)
;
;  Requires Inno Setup 6.  Run from THIS directory (repo_root\build) so the
;  relative source path  dist\Procalc\*  resolves.
; ============================================================================

#define AppName        "Procalc Hydraulics"
#define AppExeName      "Procalc.exe"
#define AppPublisher    "T.EN"
#define AppVersion      "0.1.0"
; Stable GUID identifying this application to Windows Add/Remove Programs and
; the uninstaller.  Keep it constant across versions so upgrades replace in
; place rather than installing side-by-side.
#define AppId           "{{B7B4B9C2-9E3D-4E9A-9A5B-3C1F0C7A5E10}"

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
; Install into  C:\Program Files\Procalc  ({autopf} = Program Files, matching
; the installer's bitness -- use a 64-bit build for a 64-bit exe).
DefaultDirName={autopf}\Procalc
DefaultGroupName=Procalc
; Per-machine install writing under Program Files needs admin elevation.
PrivilegesRequired=admin
; A 64-bit PyInstaller build produces a 64-bit exe: install as 64-bit so it
; lands in the real Program Files (not the WOW64 x86 folder).
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Uninstaller is generated automatically and registered in Add/Remove Programs.
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\{#AppExeName}
; Emit the installer as  Output\ProcalcSetup-<version>.exe
OutputDir=Output
OutputBaseFilename=ProcalcSetup-{#AppVersion}
; Optional: brand the installer with the app icon if it is present.
; SetupIconFile=procalc.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
; Start-menu shortcut is always created (see [Icons]); the desktop shortcut is
; opt-in via this checkbox on the "Select Additional Tasks" wizard page.
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Recursively copy the entire one-dir PyInstaller payload.  "recursesubdirs
; createallsubdirs" preserves the _internal\ tree (Qt plugins, bundled data,
; gems_extracted.json, resources\, Hydraulics\ ...).
Source: "dist\Procalc\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; NOTE: pms.json is deliberately NOT shipped or seeded by this installer.
; On first launch the app copies the bundled gems_extracted.json seed to
; %LOCALAPPDATA%\Procalc\pms.json (see pms_classes._active_json_path).  Seeding
; it here would create a machine-wide file the per-user app cannot own/edit.

[Icons]
; Start-menu shortcut (always).
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
; Uninstall entry inside the Start-menu program group.
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
; Desktop shortcut (only when the "desktopicon" task is selected).
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; Offer to launch the app when the installer finishes.
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

; ----------------------------------------------------------------------------
;  OPTIONAL: .procalc file association (commented example).
;  Uncomment the [Registry] block below -- and the extra "assocfile" task in
;  [Tasks] -- if/when Procalc gains a native project file type it should own.
;  It registers a "Procalc.Project" ProgID, points its open verb at the exe
;  ("%1" = the double-clicked file), and maps the .procalc extension to it.
; ----------------------------------------------------------------------------
; [Tasks]
; Name: "assocfile"; Description: "Associate .procalc project files with {#AppName}"; GroupDescription: "File associations:"
;
; [Registry]
; Root: HKA; Subkey: "Software\Classes\.procalc"; ValueType: string; ValueName: ""; ValueData: "Procalc.Project"; Flags: uninsdeletevalue; Tasks: assocfile
; Root: HKA; Subkey: "Software\Classes\Procalc.Project"; ValueType: string; ValueName: ""; ValueData: "Procalc Project"; Flags: uninsdeletekey; Tasks: assocfile
; Root: HKA; Subkey: "Software\Classes\Procalc.Project\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName},0"; Tasks: assocfile
; Root: HKA; Subkey: "Software\Classes\Procalc.Project\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: assocfile
