@echo off
REM build.cmd -- Builds build\DocConvert.exe and DocConvertSetup.exe, signed
REM Requires: csc.exe (in-box .NET Framework 4.x) and Inno Setup 6 (ISCC.exe)
REM   winget install JRSoftware.InnoSetup
REM Signing (v1.3.1+): Azure Artifact Signing, account dvlce-signing, profile dvlce-public.
REM   Needs signtool.exe (Windows SDK), the Artifact Signing dlib (NuGet
REM   Microsoft.ArtifactSigning.Client) and a logged-in az CLI with the
REM   "Artifact Signing Certificate Profile Signer" role. Override paths with
REM   SIGNTOOL / DLIB. Set NOSIGN=1 to build unsigned.
REM v1.3.0 replaced IExpress + launcher.vbs: antivirus flagged that chain as a trojan.

setlocal
set "ROOT=%~dp0"
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"

set "ISCC="
for %%P in ("%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" "%ProgramFiles%\Inno Setup 6\ISCC.exe") do (
    if not defined ISCC if exist %%P set "ISCC=%%~P"
)
if not defined ISCC (
    echo   ERROR: Inno Setup 6 not found. Install: winget install JRSoftware.InnoSetup
    exit /b 1
)

if defined NOSIGN goto :nosigncheck
if not defined SIGNTOOL for /f "delims=" %%S in ('dir /b /s /o:n "%ProgramFiles(x86)%\Windows Kits\10\bin\signtool.exe" 2^>nul ^| findstr /i "\\x64\\"') do set "SIGNTOOL=%%S"
if not defined DLIB set "DLIB=G:\dev-tools\artifact-signing-client\1.0.128\bin\x64\Azure.CodeSigning.Dlib.dll"
if not exist "%SIGNTOOL%" (
    echo   ERROR: signtool.exe not found. Install the Windows SDK, set SIGNTOOL, or set NOSIGN=1.
    exit /b 1
)
if not exist "%DLIB%" (
    echo   ERROR: Artifact Signing dlib not found at %DLIB%. Set DLIB, or set NOSIGN=1.
    exit /b 1
)
REM The dlib authenticates through the az CLI.
where az >nul 2>nul || set "PATH=%ProgramFiles%\Microsoft SDKs\Azure\CLI2\wbin;%PATH%"
set SIGNARGS=sign /fd SHA256 /tr http://timestamp.acs.microsoft.com /td SHA256 /dlib "%DLIB%" /dmdf "%ROOT%signing-metadata.json"
:nosigncheck

if not exist "%ROOT%build" mkdir "%ROOT%build"
"%CSC%" /nologo /target:winexe /optimize+ /win32icon:"%ROOT%doc-convert.ico" ^
 /r:Microsoft.CSharp.dll /out:"%ROOT%build\DocConvert.exe" "%ROOT%launcher\DocConvertLauncher.cs"
if errorlevel 1 (
    echo   ERROR: launcher compile failed.
    exit /b 1
)

REM Sign a copy of the script, so the source file in git stays unsigned.
copy /y "%ROOT%doc-convert.ps1" "%ROOT%build\doc-convert.ps1" >nul

if defined NOSIGN (
    "%ISCC%" /Q "%ROOT%installer\DocConvert.iss"
    goto :isccdone
)
"%SIGNTOOL%" %SIGNARGS% "%ROOT%build\DocConvert.exe" "%ROOT%build\doc-convert.ps1"
if errorlevel 1 (
    echo   ERROR: signing failed. Check: az login, and the Signer role on dvlce-signing.
    exit /b 1
)
REM Inno Setup signs the uninstaller and the setup exe with this command.
"%ISCC%" /Q /DSIGN "/Sazsign=$q%SIGNTOOL%$q sign /fd SHA256 /tr http://timestamp.acs.microsoft.com /td SHA256 /dlib $q%DLIB%$q /dmdf $q%ROOT%signing-metadata.json$q $f" "%ROOT%installer\DocConvert.iss"
:isccdone
if errorlevel 1 (
    echo   ERROR: Inno Setup build failed.
    exit /b 1
)

echo.
echo   Built: %ROOT%DocConvertSetup.exe
echo.
