@echo off
REM build.cmd — Builds DocConvertSetup.exe from the current directory
REM Requires: iexpress.exe (built into Windows, no extra install)

set "OUT=%~dp0DocConvertSetup.exe"
set "SRC=%~dp0"
REM Strip trailing backslash from SRC for IExpress compatibility
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"
set "SED=%TEMP%\docconvert_build.sed"

REM Written with a plain echo block, matching the sibling repos. The previous
REM version shelled out to PowerShell across caret-continued lines; cmd does
REM not treat ^ as a continuation inside a quoted argument, so under `call` the
REM carets reached PowerShell literally ("The term '^' is not recognized"), no
REM SED was written, and iexpress then failed on a missing file.
(
echo [Version]
echo Class=IEXPRESS
echo SEDVersion=3
echo [Options]
echo PackagePurpose=InstallApp
echo ShowInstallProgramWindow=0
echo HideExtractAnimation=0
echo UseLongFileName=1
echo InsideCompressed=0
echo CAB_FixedSize=0
echo CAB_ResvCodeSigning=0
echo RebootMode=N
echo InstallPrompt=
echo DisplayLicense=
echo FinishMessage=
echo TargetName=%OUT%
echo FriendlyName=Doc Convert Menu Installer
echo AppLaunched=setup.cmd
echo PostInstallCmd=^<None^>
echo AdminQuietInstCmd=
echo UserQuietInstCmd=
echo SourceFiles=SourceFiles
echo [Strings]
echo FILE0="setup.cmd"
echo FILE1="setup.ps1"
echo FILE2="doc-convert.ps1"
echo FILE3="launcher.vbs"
echo FILE4="install.ps1"
echo FILE5="uninstall.ps1"
echo [SourceFiles]
echo SourceFiles0=%SRC%
echo [SourceFiles0]
echo %%FILE0%%=
echo %%FILE1%%=
echo %%FILE2%%=
echo %%FILE3%%=
echo %%FILE4%%=
echo %%FILE5%%=
) > "%SED%"

REM Record the existing build so success can be judged on the timestamp.
REM "if exist" alone is a false positive: it passes on a stale .exe from a
REM previous build, which is how a broken iexpress call went unnoticed from
REM 2026-03 to 2026-07.
set "STAMP="
if exist "%OUT%" for %%F in ("%OUT%") do set "STAMP=%%~tF"

REM Do NOT quote %SED%. IExpress takes the quotes as part of the file name and
REM fails with "Error opening the IExpress Self Extraction Directive file" --
REM silently, because /Q suppresses that dialog. Keep %TEMP% free of spaces.
iexpress.exe /N /Q %SED%
del "%SED%" >nul 2>&1

set "NEWSTAMP="
if exist "%OUT%" for %%F in ("%OUT%") do set "NEWSTAMP=%%~tF"

if not defined NEWSTAMP (
    echo.
    echo   ERROR: Build failed - no output produced.
    echo          Try running as the current user without spaces in the path.
    echo.
) else if "%NEWSTAMP%"=="%STAMP%" (
    echo.
    echo   ERROR: Build failed - %OUT% is unchanged ^(%NEWSTAMP%^).
    echo          iexpress produced nothing; the old installer is still there.
    echo.
) else (
    echo.
    echo   Built: %OUT%  ^(%NEWSTAMP%^)
    echo.
)
