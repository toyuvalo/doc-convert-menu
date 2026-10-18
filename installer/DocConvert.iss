; DocConvert.iss -- Inno Setup script for DocConvertSetup.exe (v1.3.1)
; Replaces the IExpress package. IExpress extracted to %TEMP% and ran
; setup.cmd -> powershell -ExecutionPolicy Bypass, the same chain malware
; droppers use, and antivirus flagged it. This installer runs no scripts:
; it copies files and writes the HKCU context-menu keys itself.
; Build: build.cmd (needs Inno Setup 6, ISCC.exe).

#define AppVersion "1.3.1"

[Setup]
AppId={{6F1C2B7E-4D3A-4E8B-9C1F-2A7D5E0B3C91}
AppName=Doc Convert Menu
AppVersion={#AppVersion}
AppPublisher=DVLCE
AppPublisherURL=https://webdev.dvlce.ca/doc-convert
AppSupportURL=https://github.com/toyuvalo/doc-convert-menu
VersionInfoVersion={#AppVersion}
VersionInfoDescription=Doc Convert Menu Setup
; Per-user install, no admin. Same folder as the old installer, so it upgrades in place.
PrivilegesRequired=lowest
DefaultDirName={localappdata}\DocConvertMenu
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir=..
OutputBaseFilename=DocConvertSetup
SetupIconFile=..\doc-convert.ico
UninstallDisplayIcon={app}\DocConvert.exe
UninstallDisplayName=Doc Convert Menu
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ChangesAssociations=yes
#ifdef SIGN
; build.cmd passes /DSIGN and /Sazsign=<signtool command> (Azure Artifact Signing).
SignTool=azsign
SignedUninstaller=yes
#endif

[Files]
Source: "..\build\DocConvert.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\build\doc-convert.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\doc-convert.ico"; DestDir: "{app}"; Flags: ignoreversion

[InstallDelete]
; Left behind by the IExpress-era installer.
Type: files; Name: "{app}\launcher.vbs"
Type: files; Name: "{app}\setup.cmd"
Type: files; Name: "{app}\setup.ps1"
Type: files; Name: "{app}\install.ps1"
Type: files; Name: "{app}\uninstall.ps1"

#define MenuKey(Ext) "Software\Classes\SystemFileAssociations\" + Ext + "\shell\DocConvert"
#define Entry(Ext) \
  'Root: HKCU; Subkey: "' + MenuKey(Ext) + '"; ValueType: string; ValueName: ""; ValueData: "Convert with Doc Convert"; Flags: uninsdeletekey' + NewLine + \
  'Root: HKCU; Subkey: "' + MenuKey(Ext) + '"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\doc-convert.ico,0"' + NewLine + \
  'Root: HKCU; Subkey: "' + MenuKey(Ext) + '"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Player"' + NewLine + \
  'Root: HKCU; Subkey: "' + MenuKey(Ext) + '\command"; ValueType: string; ValueName: ""; ValueData: """{app}\DocConvert.exe"" ""%1"""' + NewLine

[Registry]
#emit Entry(".jpg")
#emit Entry(".jpeg")
#emit Entry(".png")
#emit Entry(".webp")
#emit Entry(".bmp")
#emit Entry(".tiff")
#emit Entry(".tif")
#emit Entry(".gif")
#emit Entry(".heic")
#emit Entry(".heif")
#emit Entry(".avif")
#emit Entry(".jxl")
#emit Entry(".jp2")
#emit Entry(".pdf")
#emit Entry(".docx")
#emit Entry(".doc")
#emit Entry(".odt")
#emit Entry(".rtf")
#emit Entry(".xlsx")
#emit Entry(".xls")
#emit Entry(".ods")
#emit Entry(".pptx")
#emit Entry(".ppt")
#emit Entry(".odp")
