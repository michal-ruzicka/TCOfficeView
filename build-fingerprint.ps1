<#
.SYNOPSIS
    Describes one build.cmd build (one bitness) so that two builds can be compared line by line.

.DESCRIPTION
    Called by build.cmd after each bitness is built. Writes, in a stable order and with
    machine-specific folders replaced by placeholders (<VS>, <VCTools>, <WindowsSdkDir>, <src>,
    <build>):

      - the toolchain: Visual Studio, MSVC toolset and Windows SDK versions, and version and SHA-256
        of cl.exe (with c1xx.dll and c2.dll), link.exe, lib.exe, rc.exe, cvtres.exe, mt.exe, cmake.exe;
      - SHA-256 of the libraries the plugin is linked with, as the linker finds them in LIB;
      - the compiler, resource compiler and linker options CMake generated;
      - SHA-256 of every object and resource file of the plugin;
      - SHA-256 of the plugin DLL and the host EXE, their Rich headers (the tool build of everything
        linked into them) and SHA-256 of each PE section.

    When two builds of the same commit produce different binaries, the first line that differs between
    their fingerprints shows where the difference comes from (see CONTRIBUTING.md, "When the hashes
    differ"). The file never goes into the release ZIP.

.PARAMETER BuildDir
    The CMake build folder of one bitness (build\x64 or build\x86).

.PARAMETER SourceDir
    The src folder (only used to hide its location).

.PARAMETER OutFile
    Where to write the fingerprint; it is printed as well.
#>
param(
    [Parameter(Mandatory = $true)][string]$BuildDir,
    [Parameter(Mandatory = $true)][string]$SourceDir,
    [Parameter(Mandatory = $true)][string]$OutFile
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$lines = New-Object System.Collections.Generic.List[string]
function Add-Line([string]$Text) { $lines.Add($Text) }

function Get-Sha256([string]$Path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try { return [System.BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '').ToLowerInvariant() }
        finally { $stream.Dispose() }
    }
    finally { $sha.Dispose() }
}

function Get-BytesSha256([byte[]]$Bytes, [int]$Offset, [int]$Count) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash($Bytes, $Offset, $Count)).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

# Machine-specific folders, longest first so that <VCTools> wins over <VS>
$placeholders = New-Object System.Collections.Generic.List[object]
function Add-Placeholder([string]$Path, [string]$Name) {
    if ([string]::IsNullOrEmpty($Path)) { return }
    $full = $Path.TrimEnd('\', '/')
    if ($full.Length -lt 3) { return }
    $placeholders.Add(@($full, $Name))
}
Add-Placeholder $env:VCToolsInstallDir '<VCTools>'
Add-Placeholder $env:WindowsSdkDir '<WindowsSdkDir>'
Add-Placeholder $env:VSINSTALLDIR '<VS>'
Add-Placeholder ([System.IO.Path]::GetFullPath($BuildDir)) '<build>'
Add-Placeholder ([System.IO.Path]::GetFullPath($SourceDir)) '<src>'
$sortedPlaceholders = @($placeholders | Sort-Object -Property @{ Expression = { $_[0].Length } } -Descending)

function Hide-Paths([string]$Text) {
    if ($null -eq $Text) { return '' }
    foreach ($p in $sortedPlaceholders) {
        foreach ($form in @($p[0], $p[0].Replace('\', '/'))) {
            $Text = [regex]::Replace($Text, [regex]::Escape($form), $p[1].Replace('$', '$$'), 'IgnoreCase')
        }
    }
    return $Text
}

function Get-FileLine([string]$Label, [string]$Path) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return ('{0,-22} (not found)' -f $Label) }
    $version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path).FileVersion
    if (-not $version) { $version = '-' }
    return ('{0,-22} {1}  {2}  {3}' -f $Label, (Get-Sha256 $Path), $version, (Hide-Paths $Path))
}

function Find-Tool([string]$Name) {
    $cmd = Get-Command $Name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) { return $cmd.Source }
    return $null
}

$targets = @('TCOfficeView', 'TCOfficeViewHost')
$binaries = @('TCOfficeView.wlx*', 'TCOfficeViewHost.exe')

# --- Toolchain --------------------------------------------------------------------------------
Add-Line '[toolchain]'
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$vsVersion = '(vswhere not found)'
if (Test-Path -LiteralPath $vswhere -PathType Leaf) {
    $vsVersion = (& $vswhere -latest -products * -property installationVersion | Select-Object -First 1)
    $vsVersion = '{0} ({1})' -f $vsVersion, (& $vswhere -latest -products * -property productId | Select-Object -First 1)
}
Add-Line ('Visual Studio          {0}, VsDevCmd {1}' -f $vsVersion, $env:VSCMD_VER)
Add-Line ('MSVC toolset           {0}' -f $env:VCToolsVersion)
$sdkProduct = '-'
foreach ($key in @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SDKs\Windows\v10.0',
                   'HKLM:\SOFTWARE\Microsoft\Microsoft SDKs\Windows\v10.0')) {
    $item = Get-ItemProperty -Path $key -Name ProductVersion -ErrorAction SilentlyContinue
    if ($item) { $sdkProduct = $item.ProductVersion; break }
}
Add-Line ('Windows SDK            {0} (UCRT {1}; newest installed SDK: {2})' -f `
    "$env:WindowsSDKVersion".TrimEnd('\'), $env:UCRTVersion, $sdkProduct)
Add-Line ('Target architecture    {0}' -f $env:VSCMD_ARG_TGT_ARCH)

$cl = Find-Tool 'cl.exe'
Add-Line (Get-FileLine 'cl.exe' $cl)
if ($cl) {
    $clDir = Split-Path -Parent $cl
    Add-Line (Get-FileLine 'c1xx.dll' (Join-Path $clDir 'c1xx.dll'))
    Add-Line (Get-FileLine 'c2.dll' (Join-Path $clDir 'c2.dll'))
}
foreach ($tool in @('link.exe', 'lib.exe', 'rc.exe', 'cvtres.exe', 'mt.exe', 'cmake.exe')) {
    Add-Line (Get-FileLine $tool (Find-Tool $tool))
}
Add-Line ('{0,-22} {1} (informational)' -f 'Windows', [System.Environment]::OSVersion.Version)
Add-Line ('{0,-22} {1} (packs the ZIP)' -f 'PowerShell', $PSVersionTable.PSVersion)

# --- Libraries, in the order the linker searches LIB -----------------------------------------
Add-Line ''
Add-Line '[libraries]'
$libDirs = @(($env:LIB -split ';') | Where-Object { $_ -ne '' })
foreach ($lib in @('libcmt.lib', 'libcpmt.lib', 'libvcruntime.lib', 'libucrt.lib', 'oldnames.lib', 'kernel32.lib', 'user32.lib', 'gdi32.lib', 'winspool.lib', 'shell32.lib', 'shlwapi.lib', 'ole32.lib', 'oleaut32.lib', 'uuid.lib', 'comdlg32.lib', 'advapi32.lib')) {
    $found = $null
    foreach ($dir in $libDirs) {
        $candidate = Join-Path $dir $lib
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $found = $candidate; break }
    }
    Add-Line (Get-FileLine $lib $found)
}

# --- Options CMake generated ------------------------------------------------------------------
Add-Line ''
Add-Line '[options]'
foreach ($target in $targets) {
    $flags = Join-Path $BuildDir "CMakeFiles\$target.dir\flags.make"
    if (Test-Path -LiteralPath $flags) {
        foreach ($line in [System.IO.File]::ReadAllLines($flags)) {
            if ($line -match '^(CXX|RC)_(FLAGS|DEFINES|INCLUDES) = ') { Add-Line ('{0}: {1}' -f $target, (Hide-Paths $line)) }
        }
    }
    # The NMake generator writes the link options into an inline response file in build.make
    $buildMake = Join-Path $BuildDir "CMakeFiles\$target.dir\build.make"
    if (Test-Path -LiteralPath $buildMake) {
        $inLink = $false
        foreach ($line in [System.IO.File]::ReadAllLines($buildMake)) {
            if ($line -match 'vs_link_(dll|exe)') { $inLink = $true; Add-Line ('{0} link: {1}' -f $target, (Hide-Paths $line.Trim())); continue }
            if ($inLink) {
                if ($line.Trim() -eq '<<') { break }
                Add-Line ('{0} link: {1}' -f $target, (Hide-Paths $line.Trim()))
            }
        }
    }
}

# --- Objects and resources ---------------------------------------------------------------------
Add-Line ''
Add-Line '[objects]'
$objects = New-Object 'System.Collections.Generic.SortedDictionary[string,string]' ([System.StringComparer]::Ordinal)
foreach ($target in $targets) {
    $dir = Join-Path $BuildDir "CMakeFiles\$target.dir"
    if (Test-Path -LiteralPath $dir) {
        Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in '.obj', '.res' } | ForEach-Object {
            $objects["$target.dir\" + $_.FullName.Substring($dir.Length).TrimStart('\', '/')] = $_.FullName
        }
    }
}
foreach ($entry in $objects.GetEnumerator()) { Add-Line ('{0}  {1}' -f (Get-Sha256 $entry.Value), $entry.Key) }
$generated = Join-Path $BuildDir 'version.h'
if (Test-Path -LiteralPath $generated) { Add-Line ('{0}  version.h (generated)' -f (Get-Sha256 $generated)) }

# --- The shipped binaries ----------------------------------------------------------------------
function Add-BinaryFingerprint([string]$Pattern) {
    $file = Get-ChildItem -LiteralPath $BuildDir -File | Where-Object { $_.Name -like $Pattern } | Select-Object -First 1
    if (-not $file) {
        Add-Line ('[binary] {0} not found' -f $Pattern)
        return
    }
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    Add-Line ('[binary] {0}' -f $file.Name)
    Add-Line ('{0}  {1} bytes' -f (Get-BytesSha256 $bytes 0 $bytes.Length), $bytes.Length)
    $peOffset = [System.BitConverter]::ToInt32($bytes, 0x3C)

    # Rich header: XOR-masked (product id, tool build, count) entries between "DanS" and "Rich"
    $richAt = -1
    for ($i = 0x80; $i -le $peOffset - 8; $i += 4) {
        if ($bytes[$i] -eq 0x52 -and $bytes[$i + 1] -eq 0x69 -and $bytes[$i + 2] -eq 0x63 -and $bytes[$i + 3] -eq 0x68) { $richAt = $i; break }
    }
    if ($richAt -lt 0) {
        Add-Line 'rich header: none'
    }
    else {
        $key = [System.BitConverter]::ToUInt32($bytes, $richAt + 4)
        $danS = 0x80
        for ($i = 0x40; $i -lt $richAt; $i += 4) {
            if (([System.BitConverter]::ToUInt32($bytes, $i) -bxor $key) -eq 0x536E6144) { $danS = $i; break }
        }
        for ($i = $danS + 16; $i -lt $richAt; $i += 8) {
            $compId = [System.BitConverter]::ToUInt32($bytes, $i) -bxor $key
            $count = [System.BitConverter]::ToUInt32($bytes, $i + 4) -bxor $key
            Add-Line ('rich: product {0,3} build {1,5} x {2}' -f ($compId -shr 16), ($compId -band 0xFFFF), $count)
        }
    }

    $sections = [System.BitConverter]::ToUInt16($bytes, $peOffset + 6)
    $optionalSize = [System.BitConverter]::ToUInt16($bytes, $peOffset + 20)
    $optional = $peOffset + 24
    Add-Line ('linker version {0}.{1}' -f $bytes[$optional + 2], $bytes[$optional + 3])
    $table = $optional + $optionalSize
    for ($s = 0; $s -lt $sections; $s++) {
        $at = $table + 40 * $s
        $name = [System.Text.Encoding]::ASCII.GetString($bytes, $at, 8).TrimEnd([char]0)
        $rawSize = [System.BitConverter]::ToInt32($bytes, $at + 16)
        $rawPointer = [System.BitConverter]::ToInt32($bytes, $at + 20)
        Add-Line ('section {0,-8} {1}  {2} bytes' -f $name, (Get-BytesSha256 $bytes $rawPointer $rawSize), $rawSize)
    }
}
foreach ($pattern in $binaries) {
    Add-Line ''
    Add-BinaryFingerprint $pattern
}

$crlf = [string][char]13 + [string][char]10
[System.IO.File]::WriteAllText($OutFile, ($lines -join $crlf) + $crlf, (New-Object System.Text.UTF8Encoding($false)))
$lines
