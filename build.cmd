@echo off
REM build.cmd — Builds DocConvertSetup.exe from the current directory
REM Requires: iexpress.exe (built into Windows, no extra install)

set "OUT=%~dp0DocConvertSetup.exe"
set "SRC=%~dp0"
REM Strip trailing backslash from SRC for IExpress compatibility
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"
set "SED=%TEMP%\docconvert_build.sed"

powershell -NoProfile -Command " ^
  $lines = @( ^
    '[Version]','Class=IEXPRESS','SEDVersion=3', ^
    '[Options]','PackagePurpose=InstallApp','ShowInstallProgramWindow=0', ^
    'HideExtractAnimation=0','UseLongFileName=1','InsideCompressed=0', ^
    'CAB_FixedSize=0','CAB_ResvCodeSigning=0','RebootMode=N', ^
    'InstallPrompt=','DisplayLicense=','FinishMessage=', ^
    'TargetName=%OUT%','FriendlyName=Doc Convert Menu Installer', ^
    'AppLaunched=setup.cmd','PostInstallCmd=<None>', ^
    'AdminQuietInstCmd=','UserQuietInstCmd=','SourceFiles=SourceFiles', ^
    '[Strings]','FILE0=setup.cmd','FILE1=setup.ps1','FILE2=doc-convert.ps1', ^
    'FILE3=launcher.vbs','FILE4=install.ps1','FILE5=uninstall.ps1', ^
    '[SourceFiles]','SourceFiles0=%SRC%','[SourceFiles0]', ^
    '%%FILE0%%=','%%FILE1%%=','%%FILE2%%=','%%FILE3%%=','%%FILE4%%=','%%FILE5%%=' ^
  ); ^
  [System.IO.File]::WriteAllLines('%SED%', $lines, [System.Text.Encoding]::ASCII)"

iexpress.exe /N /Q "%SED%"
del "%SED%" >nul 2>&1

if exist "%OUT%" (
    echo.
    echo   Built: %OUT%
    echo.
) else (
    echo.
    echo   ERROR: Build failed. Try running as the current user without spaces in the path.
    echo.
)
