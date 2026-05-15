@echo off
REM ===========================================================================
REM Build script for the TC Office Lister plugin.
REM
REM Builds both bitnesses with CMake and copies the artifacts into dist\:
REM   - tcoffice.wlx        (32-bit plugin DLL)
REM   - tcoffice.wlx64      (64-bit plugin DLL)
REM   - officehost.exe      (64-bit host EXE)
REM   - officehost_x86.exe  (32-bit host EXE)
REM
REM Prerequisites:
REM   - Visual Studio 2022 Build Tools, workload "Desktop development with C++"
REM   - CMake 3.20+
REM ===========================================================================

setlocal enabledelayedexpansion

REM %~dp0 expands with a trailing backslash; quoting "%ROOT%\" then escapes the
REM closing quote and breaks command-line parsing for cmake. Strip it.
set ROOT=%~dp0
if "%ROOT:~-1%"=="\" set ROOT=%ROOT:~0,-1%
set BUILD=%ROOT%\build
set OUT=%ROOT%\dist

if not exist "%BUILD%" mkdir "%BUILD%"
if not exist "%OUT%"   mkdir "%OUT%"

echo.
echo === Configure + build (x64) ===
cmake -S "%ROOT%" -B "%BUILD%\x64" -A x64
if errorlevel 1 goto :error
cmake --build "%BUILD%\x64" --config Release
if errorlevel 1 goto :error
copy /Y "%BUILD%\x64\Release\tcoffice.wlx64"   "%OUT%\"                     >nul
copy /Y "%BUILD%\x64\Release\officehost.exe"   "%OUT%\officehost.exe"       >nul

echo.
echo === Configure + build (x86) ===
cmake -S "%ROOT%" -B "%BUILD%\x86" -A Win32
if errorlevel 1 goto :error
cmake --build "%BUILD%\x86" --config Release
if errorlevel 1 goto :error
copy /Y "%BUILD%\x86\Release\tcoffice.wlx"     "%OUT%\"                     >nul
copy /Y "%BUILD%\x86\Release\officehost.exe"   "%OUT%\officehost_x86.exe"   >nul

echo.
echo === Done ===
echo Output: %OUT%
dir /b "%OUT%"
goto :eof

:error
echo BUILD FAILED
exit /b 1
