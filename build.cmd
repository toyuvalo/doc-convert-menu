@echo off
REM build.cmd -- Builds build\DocConvert.exe and DocConvertSetup.exe
REM Requires: csc.exe (in-box .NET Framework 4.x) and Inno Setup 6 (ISCC.exe)
REM   winget install JRSoftware.InnoSetup
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

if not exist "%ROOT%build" mkdir "%ROOT%build"
"%CSC%" /nologo /target:winexe /optimize+ /win32icon:"%ROOT%doc-convert.ico" ^
 /r:Microsoft.CSharp.dll /out:"%ROOT%build\DocConvert.exe" "%ROOT%launcher\DocConvertLauncher.cs"
if errorlevel 1 (
    echo   ERROR: launcher compile failed.
    exit /b 1
)

"%ISCC%" /Q "%ROOT%installer\DocConvert.iss"
if errorlevel 1 (
    echo   ERROR: Inno Setup build failed.
    exit /b 1
)

echo.
echo   Built: %ROOT%DocConvertSetup.exe
echo.
