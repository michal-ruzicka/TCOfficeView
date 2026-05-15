@echo off
REM ===========================================================================
REM Build script pro TC Office Lister plugin
REM
REM Předpoklady:
REM   - Visual Studio 2022 Build Tools (s C++ workload)
REM   - .NET 8 SDK
REM   - CMake 3.20+
REM ===========================================================================

setlocal enabledelayedexpansion

set ROOT=%~dp0
set BUILD=%ROOT%build
set OUT=%ROOT%dist

if not exist "%BUILD%" mkdir "%BUILD%"
if not exist "%OUT%" mkdir "%OUT%"

echo.
echo === Building C++ plugin (x64) ===
cmake -S "%ROOT%plugin" -B "%BUILD%\plugin-x64" -A x64
if errorlevel 1 goto :error
cmake --build "%BUILD%\plugin-x64" --config Release
if errorlevel 1 goto :error
copy "%BUILD%\plugin-x64\Release\tcoffice.wlx64" "%OUT%\" >nul

echo.
echo === Building C++ plugin (x86) ===
cmake -S "%ROOT%plugin" -B "%BUILD%\plugin-x86" -A Win32
if errorlevel 1 goto :error
cmake --build "%BUILD%\plugin-x86" --config Release
if errorlevel 1 goto :error
copy "%BUILD%\plugin-x86\Release\tcoffice.wlx" "%OUT%\" >nul

echo.
echo === Building C# host (x64) ===
dotnet publish "%ROOT%host\OfficePreviewHost.csproj" ^
    -c Release -r win-x64 --self-contained false ^
    -o "%BUILD%\host-x64"
if errorlevel 1 goto :error
copy "%BUILD%\host-x64\OfficePreviewHost.exe" "%OUT%\OfficePreviewHost.exe" >nul

echo.
echo === Building C# host (x86) ===
dotnet publish "%ROOT%host\OfficePreviewHost.csproj" ^
    -c Release -r win-x86 --self-contained false ^
    -o "%BUILD%\host-x86"
if errorlevel 1 goto :error
REM Pro 32-bit TC použijeme 32-bit verzi exe - sufix _x86
copy "%BUILD%\host-x86\OfficePreviewHost.exe" "%OUT%\OfficePreviewHost_x86.exe" >nul

echo.
echo === Done ===
echo Output: %OUT%
dir /b "%OUT%"
goto :eof

:error
echo BUILD FAILED
exit /b 1
