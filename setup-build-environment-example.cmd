@echo off
REM ===========================================================================
REM Sets up a minimal build environment from scratch and produces the
REM distributable ZIP.  Intended for fresh Windows installations — virtual
REM machines, Windows Sandbox (https://learn.microsoft.com/windows/security/application-security/application-isolation/windows-sandbox/) 
REM instances, and similar clean environments.
REM
REM What it does:
REM   1. Downloads the source tree from GitHub.
REM   2. Installs Visual Studio 2026 Build Tools (C++ workload).
REM   3. Installs CMake via winget.
REM   4. Runs build.cmd to produce the distributable ZIP in dist\.
REM
REM Usage:
REM   setup-build-environment-example.cmd
REM
REM To target a different branch, set BRANCH before running:
REM   set BRANCH=main
REM   setup-build-environment-example.cmd
REM ===========================================================================
setlocal

REM ---------------------------------------------------------------------------
REM Source branch.  Set BRANCH in the environment before running to override.
REM ---------------------------------------------------------------------------
if not defined BRANCH set BRANCH=main

REM GitHub archives a branch as <repo>-<branch>/ inside the ZIP, with every
REM "/" in the branch name replaced by "-".
set BRANCH_FOLDER=%BRANCH:/=-%
set SRC_DIR=C:\TCOfficeView-%BRANCH_FOLDER%

echo.
echo === Download sources (branch: %BRANCH%) ===
curl.exe -L -o "%TEMP%\TCOfficeView-src.zip" ^
    "https://github.com/michal-ruzicka/TCOfficeView/archive/refs/heads/%BRANCH%.zip"
if errorlevel 1 (echo ERROR: source download failed & exit /b 1)
tar -xf "%TEMP%\TCOfficeView-src.zip" -C C:\
if errorlevel 1 (echo ERROR: source extraction failed & exit /b 1)

echo.
echo === Install Visual Studio Build Tools ===
echo = This step takes several minutes and may appear unresponsive.
echo = Check Task Manager to confirm the Visual Studio Installer process is still running.
echo.
curl.exe -L -o "%TEMP%\vs_buildtools.exe" https://aka.ms/vs/18/stable/vs_buildtools.exe
if errorlevel 1 (echo ERROR: Build Tools installer download failed & exit /b 1)
"%TEMP%\vs_buildtools.exe" --config "%SRC_DIR%\.vsconfig" ^
    --installPath C:\BuildTools ^
    --quiet --wait --norestart --nocache
REM The VS bootstrapper returns 3010 when a reboot is requested; that is still
REM a successful install, so accept it alongside 0.
if errorlevel 1 if not errorlevel 3010 (echo ERROR: Build Tools installation failed & exit /b 1)
if %errorlevel% gtr 3010 (echo ERROR: Build Tools installation failed & exit /b 1)

REM CMake (and Ninja) ship with the Build Tools via the VC.CMake.Project
REM component in .vsconfig; build.cmd puts them on PATH through VsDevCmd.bat,
REM so no separate CMake install is needed.

echo.
echo === Build ===
cd /d "%SRC_DIR%"
build.cmd
