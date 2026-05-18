@echo off
REM ===========================================================================
REM Build script for the TCOfficeView Lister plugin.
REM
REM Builds both bitnesses with CMake, stages the artifacts under build\stage\,
REM and packages them into a Total Commander auto-install ZIP whose name
REM carries the version read from pluginst.inf. After the build, dist\
REM contains nothing but the ZIP(s) — older versioned ZIPs are preserved.
REM
REM Layout after a successful build:
REM
REM   build\stage\
REM     TCOfficeView.wlx              (32-bit plugin DLL)
REM     TCOfficeView.wlx64            (64-bit plugin DLL)
REM     TCOfficeViewHost.exe          (64-bit host EXE)
REM     TCOfficeViewHost_x86.exe      (32-bit host EXE)
REM     pluginst.inf                  (TC auto-installer metadata)
REM     TCOfficeView.ini              (system-wide sample config)
REM     *.md                          (all repo-root Markdown docs)
REM   dist\
REM     TCOfficeView.v<version>.zip   (distributable bundle)
REM     ...                           (any older versioned ZIPs are kept)
REM
REM Prerequisites:
REM   - Visual Studio 2022 Build Tools, workload "Desktop development with C++"
REM   - CMake 3.20+
REM   - PowerShell 5.1+ (ships with Windows; used for Compress-Archive)
REM ===========================================================================

setlocal enabledelayedexpansion

REM %~dp0 expands with a trailing backslash; quoting "%ROOT%\" then escapes the
REM closing quote and breaks command-line parsing for cmake. Strip it.
set ROOT=%~dp0
if "%ROOT:~-1%"=="\" set ROOT=%ROOT:~0,-1%
set BUILD=%ROOT%\build
set STAGE=%BUILD%\stage
set OUT=%ROOT%\dist

if not exist "%BUILD%" mkdir "%BUILD%"
if not exist "%OUT%"   mkdir "%OUT%"

echo.
echo === Configure + build (x64) ===
cmake -S "%ROOT%" -B "%BUILD%\x64" -A x64
if errorlevel 1 goto :error
cmake --build "%BUILD%\x64" --config Release
if errorlevel 1 goto :error

echo.
echo === Configure + build (x86) ===
cmake -S "%ROOT%" -B "%BUILD%\x86" -A Win32
if errorlevel 1 goto :error
cmake --build "%BUILD%\x86" --config Release
if errorlevel 1 goto :error

echo.
echo === Stage ===
if exist "%STAGE%" rmdir /S /Q "%STAGE%"
mkdir "%STAGE%"

copy /Y "%BUILD%\x64\Release\TCOfficeView.wlx64"        "%STAGE%\"                              >nul
copy /Y "%BUILD%\x64\Release\TCOfficeViewHost.exe"      "%STAGE%\TCOfficeViewHost.exe"          >nul
copy /Y "%BUILD%\x86\Release\TCOfficeView.wlx"          "%STAGE%\"                              >nul
copy /Y "%BUILD%\x86\Release\TCOfficeViewHost.exe"      "%STAGE%\TCOfficeViewHost_x86.exe"      >nul
copy /Y "%ROOT%\pluginst.inf"                           "%STAGE%\"                              >nul
copy /Y "%ROOT%\TCOfficeView.ini"                       "%STAGE%\"                              >nul
copy /Y "%ROOT%\*.md"                                   "%STAGE%\"                              >nul

echo.
echo === Package ===

REM Pull the version out of pluginst.inf (single source of truth). The line
REM looks like   version=0.1.0   — split on '=' and grab the second token.
set VERSION=
for /f "usebackq tokens=2 delims==" %%v in (`findstr /b /c:"version=" "%ROOT%\pluginst.inf"`) do set VERSION=%%v
if "%VERSION%"=="" (
    echo Could not read version= from pluginst.inf
    goto :error
)

set ZIP=%OUT%\TCOfficeView.v%VERSION%.zip
echo Building %ZIP%
powershell -NoLogo -NoProfile -Command ^
    "Compress-Archive -Force -DestinationPath '%ZIP%' -Path '%STAGE%\*'"
if errorlevel 1 goto :error

echo.
echo === Done ===
echo Output: %OUT%
dir /b "%OUT%"
goto :eof

:error
echo BUILD FAILED
exit /b 1
