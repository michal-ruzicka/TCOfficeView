@echo off
REM ===========================================================================
REM Sets up a minimal build environment from scratch and produces the
REM distributable ZIP.  Intended for fresh Windows installations — virtual
REM machines, Windows Sandbox (https://learn.microsoft.com/windows/security/application-security/application-isolation/windows-sandbox/) 
REM instances, and similar clean environments.
REM
REM What it does:
REM   1. Downloads the source tree from GitHub.
REM   2. Installs Visual Studio 2026 Build Tools of the release the build is
REM      pinned to (EXPECTED_VS in build.cmd, VS_BUILDTOOLS_* in the CI
REM      workflow) with the components listed in .vsconfig (C++ toolset,
REM      Windows SDK, CMake).
REM   3. Runs build.cmd to produce the distributable ZIP in dist\.
REM
REM Usage:
REM   setup-build-environment-example.cmd
REM
REM To target a different branch, set BRANCH before running:
REM   set BRANCH=my-feature-branch
REM   setup-build-environment-example.cmd
REM ===========================================================================
setlocal

REM ---------------------------------------------------------------------------
REM Source branch.  Set BRANCH in the environment before running to override.
REM ---------------------------------------------------------------------------
if not defined BRANCH set BRANCH=main

REM ---------------------------------------------------------------------------
REM Visual Studio Build Tools release: the fixed-version bootstrapper and its
REM SHA-256, the same as VS_BUILDTOOLS_URL / VS_BUILDTOOLS_SHA256 in
REM .github\workflows\build.yml. Keep the two files in step.
REM ---------------------------------------------------------------------------
set VS_BUILDTOOLS_URL=https://download.visualstudio.microsoft.com/download/pr/4aa37a90-1dc9-4507-9e7b-49134341b6c4/5e7c32e82c89b3d69a2c818f3b06679909729b05e94609e3367bbb554d627f55/vs_BuildTools.exe
set VS_BUILDTOOLS_SHA256=5e7c32e82c89b3d69a2c818f3b06679909729b05e94609e3367bbb554d627f55

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
curl.exe -L -o "%TEMP%\vs_buildtools.exe" %VS_BUILDTOOLS_URL%
if errorlevel 1 (echo ERROR: Build Tools installer download failed & exit /b 1)
powershell -NoLogo -NoProfile -Command ^
    "if ((Get-FileHash -Algorithm SHA256 '%TEMP%\vs_buildtools.exe').Hash -ne '%VS_BUILDTOOLS_SHA256%') { exit 1 }"
if errorlevel 1 (echo ERROR: Build Tools installer does not match its SHA-256 & exit /b 1)
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
REM Use the installation made above even if the machine has other ones.
set TCOV_VSINSTALL=C:\BuildTools
cd /d "%SRC_DIR%"
build.cmd
