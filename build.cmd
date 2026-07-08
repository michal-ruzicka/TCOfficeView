@echo off
REM ===========================================================================
REM Build script for the TCOfficeView Lister plugin.
REM
REM Reads sources from src\, builds both bitnesses with CMake, stages the
REM artifacts under build\stage\, and packages them into a Total Commander
REM auto-install ZIP whose name carries the version read from
REM src\pluginst.inf. After the build, dist\ contains nothing but the ZIP(s) —
REM older versioned ZIPs are preserved.
REM
REM Layout after a successful build:
REM
REM   src\                              (all source files — committed)
REM   build\stage\                      (transient staging for ZIP contents)
REM     TCOfficeView.wlx                (32-bit plugin DLL)
REM     TCOfficeView.wlx64              (64-bit plugin DLL)
REM     TCOfficeViewHost.exe            (64-bit host EXE)
REM     TCOfficeViewHost_x86.exe        (32-bit host EXE)
REM     sha256sums.sha256               (SHA-256 of the four binaries above)
REM     pluginst.inf                    (TC auto-installer metadata)
REM     TCOfficeView.ini                (system-wide sample config)
REM     *.md                            (all repo-root Markdown docs)
REM   dist\
REM     TCOfficeView.v<version>.zip     (distributable bundle)
REM     ...                             (any older versioned ZIPs are kept)
REM
REM Prerequisites:
REM   - Visual Studio 2026 Build Tools, workload "Desktop development with C++"
REM   - CMake 3.20+ (NMake ships with VS Build Tools; no separate install needed)
REM   - PowerShell 5.1+ (ships with Windows; used for Compress-Archive)
REM ===========================================================================

REM ---------------------------------------------------------------------------
REM Pinned toolchain versions — update all three when upgrading (see
REM CONTRIBUTING.md, section "Reproducible Builds / Upgrading the pinned
REM toolchain"):
REM   EXPECTED_CL  — exact cl.exe version reported by "cl.exe /Bv"
REM   CMAKE_T      — -vcvars_ver= token passed to VsDevCmd.bat (toolset major.minor)
REM   CMAKE_SDK    — -DCMAKE_SYSTEM_VERSION= value (Windows SDK)
REM ---------------------------------------------------------------------------
set EXPECTED_CL=19.51.36248
set CMAKE_T=14.51
set CMAKE_SDK=10.0.26100.0

setlocal enabledelayedexpansion

REM %~dp0 expands with a trailing backslash; quoting "%ROOT%\" then escapes the
REM closing quote and breaks command-line parsing for cmake. Strip it.
set ROOT=%~dp0
if "%ROOT:~-1%"=="\" set ROOT=%ROOT:~0,-1%
set SRC=%ROOT%\src
set BUILD=%ROOT%\build
set STAGE=%BUILD%\stage
set OUT=%ROOT%\dist

if not exist "%BUILD%" mkdir "%BUILD%"
if not exist "%OUT%"   mkdir "%OUT%"

REM Locate Build Tools via vswhere and activate the x64 compiler environment.
REM cmake's VS generator instance detection fails for Build Tools installed
REM without VS IDE, so we use the Ninja generator and configure the compiler
REM environment explicitly via VsDevCmd.bat for each target architecture.
for /f "usebackq tokens=*" %%i in (`"C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe" -products * -latest -property installationPath`) do set VSINSTALL=%%i
if not defined VSINSTALL (
    echo ERROR: Visual Studio Build Tools not found. Run setup-build-environment.cmd first.
    goto :error
)
set VSDEVCMD=%VSINSTALL%\Common7\Tools\VsDevCmd.bat
call "%VSDEVCMD%" -startdir=none -arch=x64 -host_arch=x64 -vcvars_ver=%CMAKE_T%
if errorlevel 1 goto :error

echo.
echo === Verify compiler version ===
REM Print the cl.exe version so it is visible in the build log. If a local
REM build and a CI build produce different binary hashes, compare this line
REM in both logs — a version mismatch is the first thing to check.
REM
REM A mismatch is a WARNING, not an error: cl.exe patch versions advance
REM independently on local machines and CI runner images and cannot be
REM synchronised without coordinated updates on both sides. EXPECTED_CL
REM documents the version at the time the pin was last set; update it
REM (along with a matching VS update on the other side) when you want to
REM re-establish bit-for-bit parity.
for /f "tokens=7 delims= " %%v in ('cl.exe 2^>^&1 ^| findstr /C:"Compiler Version"') do set ACTUAL_CL=%%v
set CL_MISMATCH=0
if not defined ACTUAL_CL (
    echo WARNING: could not determine cl.exe version
    set CL_MISMATCH=1
) else if not "%ACTUAL_CL%"=="%EXPECTED_CL%" (
    echo WARNING: cl.exe version mismatch
    echo   expected : %EXPECTED_CL%
    echo   actual   : %ACTUAL_CL%
    echo   Binaries from this build may differ from builds using %EXPECTED_CL%.
    set CL_MISMATCH=1
) else (
    echo cl.exe %ACTUAL_CL% OK
)

echo.
echo === Configure + build (x64) ===
cmake -S "%SRC%" -B "%BUILD%\x64" -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release -DCMAKE_SYSTEM_VERSION=%CMAKE_SDK%
if errorlevel 1 goto :error
cmake --build "%BUILD%\x64"
if errorlevel 1 goto :error

echo.
echo === Configure + build (x86) ===
call "%VSDEVCMD%" -startdir=none -arch=x86 -host_arch=x64 -vcvars_ver=%CMAKE_T%
if errorlevel 1 goto :error
cmake -S "%SRC%" -B "%BUILD%\x86" -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release -DCMAKE_SYSTEM_VERSION=%CMAKE_SDK%
if errorlevel 1 goto :error
cmake --build "%BUILD%\x86"
if errorlevel 1 goto :error

echo.
echo === Stage ===
if exist "%STAGE%" rmdir /S /Q "%STAGE%"
mkdir "%STAGE%"

copy /Y "%BUILD%\x64\TCOfficeView.wlx64"             "%STAGE%\"                              >nul
copy /Y "%BUILD%\x64\TCOfficeViewHost.exe"           "%STAGE%\TCOfficeViewHost.exe"          >nul
copy /Y "%BUILD%\x86\TCOfficeView.wlx"               "%STAGE%\"                              >nul
copy /Y "%BUILD%\x86\TCOfficeViewHost.exe"           "%STAGE%\TCOfficeViewHost_x86.exe"      >nul
copy /Y "%SRC%\pluginst.inf"                         "%STAGE%\"                              >nul
copy /Y "%SRC%\TCOfficeView.ini"                     "%STAGE%\"                              >nul
copy /Y "%ROOT%\*.md"                                "%STAGE%\"                              >nul

echo.
echo === Normalize line endings (staged text files) ===
REM Rewrite all line endings to CRLF in every text file in the stage.
REM Must run before the checksum step so hashes cover the final content.
REM This is a Windows-only tool; CRLF is the native line ending and ensures
REM text files open correctly in any Windows editor.
REM Uses [char] casts to avoid batch quoting issues with escape sequences.
powershell -NoLogo -NoProfile -Command ^
    "Get-ChildItem '%STAGE%' -File | Where-Object { $_.Extension -in '.md','.inf','.ini','.txt','.sha256' } | ForEach-Object { $lf = [string][char]10; $cr = [string][char]13; $crlf = $cr + $lf; $t = [System.IO.File]::ReadAllText($_.FullName); [System.IO.File]::WriteAllText($_.FullName, $t.Replace($crlf, $lf).Replace($cr, $lf).Replace($lf, $crlf), [System.Text.UTF8Encoding]::new($false)) }"
if errorlevel 1 goto :error

echo.
echo === Checksums (all staged files) ===
REM SHA-256 of every file in the stage except sha256sums.sha256 itself, sorted
REM by name for a stable order. Written with CRLF so it opens correctly in
REM any Windows editor. Users can verify extracted files against a known-good
REM build; CI log shows the same hashes for cross-build comparison.
powershell -NoLogo -NoProfile -Command ^
    "$crlf = [string][char]13 + [string][char]10; $h = [System.Security.Cryptography.SHA256]::Create(); $d = [System.Collections.Generic.SortedDictionary[string,System.IO.FileInfo]]::new([System.StringComparer]::Ordinal); Get-ChildItem '%STAGE%' -File | Where-Object { $_.Name -ne 'sha256sums.sha256' } | ForEach-Object { $d[$_.Name] = $_ }; $lines = $d.Values | ForEach-Object { [System.BitConverter]::ToString($h.ComputeHash([System.IO.File]::ReadAllBytes($_.FullName))).Replace('-','').ToLower() + '  ' + $_.Name }; [System.IO.File]::WriteAllText('%STAGE%\sha256sums.sha256', ($lines -join $crlf) + $crlf, [System.Text.UTF8Encoding]::new($false)); $lines"
if errorlevel 1 goto :error

echo.
echo === Normalize file timestamps ===
REM Set all staged file timestamps to the release date from pluginst.inf so
REM that ZIP entries have a deterministic mtime regardless of when or where
REM the build runs (no git dependency).
set RELEASE_DATE=
for /f "usebackq tokens=2 delims==" %%d in (`findstr /b /c:"release-date=" "%SRC%\pluginst.inf"`) do set RELEASE_DATE=%%d
if "%RELEASE_DATE%"=="" (
    echo Could not read release-date= from pluginst.inf
    goto :error
)
powershell -NoLogo -NoProfile -Command ^
    "$t = [datetime]::Parse('%RELEASE_DATE%'); Get-ChildItem -Path '%STAGE%' | ForEach-Object { $_.LastWriteTime = $t; $_.CreationTime = $t }"
if errorlevel 1 goto :error

echo.
echo === Package ===

REM Pull the version out of pluginst.inf (single source of truth). The line
REM looks like   version=0.1.0   — split on '=' and grab the second token.
set VERSION=
for /f "usebackq tokens=2 delims==" %%v in (`findstr /b /c:"version=" "%SRC%\pluginst.inf"`) do set VERSION=%%v
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
echo.
echo SHA-256 (ZIP):
powershell -NoLogo -NoProfile -Command ^
    "$h = [System.Security.Cryptography.SHA256]::Create(); [System.BitConverter]::ToString($h.ComputeHash([System.IO.File]::ReadAllBytes('%ZIP%'))).Replace('-','').ToLower() + '  ' + (Split-Path '%ZIP%' -Leaf)"
if %CL_MISMATCH%==1 (
    echo.
    if not defined ACTUAL_CL (
        echo WARNING: cl.exe version could not be determined; reproducible build not achieved.
    ) else (
        echo WARNING: cl.exe version mismatch ^(expected %EXPECTED_CL%, used %ACTUAL_CL%^); binaries from this build can differ from what was expected. Reproducible build not achieved.
    )
)
goto :eof

:error
echo BUILD FAILED
exit /b 1
