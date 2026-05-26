// SPDX-License-Identifier: Apache-2.0
// Copyright 2026 Michal Růžička <ruzicka.mich@gmail.com>

/*
 * TCOfficeViewHost.cpp - Total Commander Office preview host process.
 *
 * Spawned by TCOfficeView.wlx. Hosts a Windows Preview Handler COM object
 * and renders it inside Total Commander's Lister pane.
 *
 * Window topology:
 *
 *      Total Commander process              this host process
 *      ---------------------                ------------------
 *      Lister parent
 *        └── plugin child   (TCOfficeView DLL)
 *              └── (SetParent) ─────────── render window  (this process)
 *                                              └── preview handler's
 *                                                  internal windows
 *
 *   The DLL creates a child window inside TC's Lister pane and passes its
 *   HWND on our command line. We then create our *own* child window in *this*
 *   process and SetParent it into the DLL's window. The preview handler is
 *   given our render window — never TC's — so its internal SetParent calls
 *   stay within our process and don't deadlock against TC's UI thread.
 *
 * Threading model:
 *   - Main thread is the STA. Runs a Win32 message loop. All COM calls and
 *     window-tree manipulation happen here.
 *   - Pipe reader runs on a worker thread, dispatches each command to the
 *     STA window via SendMessage, and writes the response back.
 *
 * Wire protocol (UTF-16 LE, one message per WriteFile, message-mode pipe):
 *   LOAD <absolute-path>\n
 *   RESIZE <width> <height>\n
 *   CLOSE\n
 * Responses written back to the plugin:
 *   OK\n
 *   ERR <message>\n
 */

#include <windows.h>
#include <winioctl.h>         // FSCTL_SET_REPARSE_POINT for the long-path junction helper
#include <shobjidl.h>
#include <shlwapi.h>
#include <shlobj.h>           // SHCreateDirectoryExW
#include <propsys.h>          // IInitializeWithFile, IInitializeWithStream, IInitializeWithItem
#include <ocidl.h>            // IObjectWithSite
#include <servprov.h>         // IServiceProvider
#include <objbase.h>
#include <oleauto.h>          // IDispatch helpers: SysAllocString, VARIANT, ...
#include <stdio.h>
#include <string>
#include <vector>
#include <map>             // [PreviewHandlers] override table
#include <set>             // [PreviewHandlers] deny list
#include <algorithm>       // std::sort in the available-handlers report

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")      // UpdateLayeredWindow, SetLayeredWindowAttributes

// ---------------------------------------------------------------------------
// Runtime configuration (loaded from TCOfficeView.ini)
//
// Lookup order (first existing file wins; no per-key merging):
//   1. %APPDATA%\GHISLER\TCOfficeView.ini       — per-user override
//   2. <plugin install dir>\TCOfficeView.ini    — system-wide default
//
// All values support environment-variable expansion (e.g. %TEMP%,
// %LocalAppData%). See the shipped sample INI for the full key reference.
// ---------------------------------------------------------------------------

// Per-application render mode.
//
//   Quick          — Preview Handler only; mode-switch button hidden.
//   QuickSwitchable— Preview Handler by default; button visible so the
//                    user can flip to Full for a single preview.
//   Full           — OLE Automation (real app) only; button hidden.
//   FullSwitchable — OLE Automation by default; button visible so the
//                    user can flip to Quick for a single preview.
//
// BaseMode() strips the Switchable suffix to get the actual loader mode.
// IsSwitchable() tells whether the mode-switch button should be shown.
enum class Mode { Quick, Full, QuickSwitchable, FullSwitchable };

static Mode BaseMode(Mode m)
{
    return (m == Mode::Full || m == Mode::FullSwitchable) ? Mode::Full : Mode::Quick;
}
static bool IsSwitchable(Mode m)
{
    return m == Mode::QuickSwitchable || m == Mode::FullSwitchable;
}

static std::wstring g_logPath;       // empty → diagnostic logging is disabled
static std::wstring g_fontFamily;    // empty → auto-pick from a fallback list
static int          g_fontSize = 12;
static Mode         g_modeWord       = Mode::QuickSwitchable;   // default: quick + button
static Mode         g_modeExcel      = Mode::QuickSwitchable;
static Mode         g_modePowerPoint = Mode::QuickSwitchable;

// Per-app auto-fallback toggles for the quick-switchable mode.  When the
// configured mode is QuickSwitchable AND auto-fallback is enabled, a quick-
// mode load failure for an Office file silently retries in full mode.  The
// typical cause we are guarding against is SharePoint documents synced from
// non-primary Microsoft 365 tenants: the prevhost.exe surrogate the Office
// preview handlers run in cannot authenticate cross-tenant, so quick-mode
// rendering fails with E_FAIL / E_NOTIMPL.  Spinning up the real Office app
// (which runs as the user with full auth tokens) gets the preview through.
// Defaults to true for all three apps; users who prefer the explicit
// fallback panel over Office's slower cold-start can flip the toggle off
// in [AutoFallback] in TCOfficeView.ini.
static bool         g_autoFallbackWord       = true;
static bool         g_autoFallbackExcel      = true;
static bool         g_autoFallbackPowerPoint = true;

// Per-extension preview-handler CLSID overrides from the INI's
// [PreviewHandlers] section.  Keys are lowercase extensions with the
// leading dot ("`.pdf`"), values are the raw CLSID string in braces
// ("`{1531D583-...}`").  Consulted first by FindPreviewHandlerClsid;
// the system registry chain is the fallback when no override matches.
//
// Use case: multiple PDF preview handlers are installed (Edge plus
// Adobe Reader, for example) and the user wants this plugin to pick
// one regardless of which one Windows considers the system default.
static std::map<std::wstring, std::wstring> g_extensionOverrides;

// Per-extension deny list from [PreviewHandlers] in the INI: an entry
// of the form ".pdf=" (empty value) marks the extension as
// "intentionally unhandled by this plugin".  FindPreviewHandlerClsid
// returns REGDB_E_CLASSNOTREG for such extensions so they behave the
// same as files for which no preview handler is registered at all —
// the plugin DLL declines and Total Commander tries the next plugin.
// The plugin DLL maintains its own parallel deny set so it can avoid
// spawning the host process entirely; this set in the host is a
// defence-in-depth backstop.
static std::set<std::wstring> g_deniedExtensions;

// Optional opt-in discovery aid (see [PreviewHandlers] ReportPath in
// the INI).  When non-empty, the host writes a human-readable list of
// every installed preview handler and every per-extension assignment
// to this path on each start-up.  Off by default — most users only
// ever need it once, when first setting up an override.
static std::wstring g_handlersReportPath;

// Forward declaration — HostLog's real definition lives further down,
// inside the #if HOST_LOG_ENABLED block.  Declared here so the
// LoadConfigFrom parser below can announce loaded / rejected
// [PreviewHandlers] overrides.
static void HostLog(const wchar_t* fmt, ...);

static std::wstring ExpandEnv(LPCWSTR s)
{
    if (!s || !*s) return L"";
    DWORD n = ExpandEnvironmentStringsW(s, nullptr, 0);
    if (n == 0) return s;
    std::wstring out(n, L'\0');
    DWORD written = ExpandEnvironmentStringsW(s, out.data(), n);
    if (written == 0) return s;
    out.resize(written ? written - 1 : 0);   // drop the trailing null
    return out;
}

static std::wstring GetHostExeDir()
{
    // Grow the buffer until GetModuleFileNameW reports the full path fits.
    // The host EXE may sit under a path longer than MAX_PATH; with a fixed-
    // size buffer we'd silently truncate and the system-wide INI lookup
    // would target the wrong directory.
    std::wstring path;
    DWORD bufSize = MAX_PATH;
    for (;;)
    {
        path.resize(bufSize);
        DWORD len = GetModuleFileNameW(nullptr, path.data(), bufSize);
        if (len == 0) return L".";
        if (len < bufSize) { path.resize(len); break; }
        bufSize *= 2;
        if (bufSize > 32768) return L".";
    }
    auto pos = path.find_last_of(L"\\/");
    return (pos != std::wstring::npos) ? path.substr(0, pos) : L".";
}

static bool LoadConfigFrom(const std::wstring& iniPath)
{
    if (GetFileAttributesW(iniPath.c_str()) == INVALID_FILE_ATTRIBUTES)
        return false;
    wchar_t buf[2048] = {};

    GetPrivateProfileStringW(L"Logging", L"LogPath", L"",
                             buf, ARRAYSIZE(buf), iniPath.c_str());
    g_logPath = ExpandEnv(buf);

    GetPrivateProfileStringW(L"FallbackUI", L"FontFamily", L"",
                             buf, ARRAYSIZE(buf), iniPath.c_str());
    g_fontFamily = buf;

    g_fontSize = GetPrivateProfileIntW(L"FallbackUI", L"FontSize",
                                       12, iniPath.c_str());
    if (g_fontSize < 6)  g_fontSize = 6;     // clamp to a sane range
    if (g_fontSize > 72) g_fontSize = 72;

    auto readMode = [&buf, &iniPath](LPCWSTR key, Mode dflt) -> Mode {
        GetPrivateProfileStringW(L"Mode", key, L"",
                                 buf, ARRAYSIZE(buf), iniPath.c_str());
        if (_wcsicmp(buf, L"quick")            == 0) return Mode::Quick;
        if (_wcsicmp(buf, L"quick-switchable") == 0) return Mode::QuickSwitchable;
        if (_wcsicmp(buf, L"full")             == 0) return Mode::Full;
        if (_wcsicmp(buf, L"full-switchable")  == 0) return Mode::FullSwitchable;
        return dflt;                          // empty or invalid → default
    };
    g_modeWord       = readMode(L"Word",       Mode::QuickSwitchable);
    g_modeExcel      = readMode(L"Excel",      Mode::QuickSwitchable);
    g_modePowerPoint = readMode(L"PowerPoint", Mode::QuickSwitchable);

    // [AutoFallback] — per-app quick→full auto-fallback toggles.  Only
    // consulted when the configured Mode is `quick-switchable` (the user
    // has opted into mode-switching for that app, so silently flipping
    // modes on a per-file basis matches the spirit of that setting).
    // Accepted values are true/false, yes/no, on/off, 1/0 — anything else
    // (including missing) keeps the default of `true`.
    auto readBool = [&buf, &iniPath](LPCWSTR section, LPCWSTR key, bool dflt) -> bool {
        GetPrivateProfileStringW(section, key, L"",
                                 buf, ARRAYSIZE(buf), iniPath.c_str());
        if (buf[0] == L'\0') return dflt;
        if (_wcsicmp(buf, L"true")  == 0) return true;
        if (_wcsicmp(buf, L"false") == 0) return false;
        if (_wcsicmp(buf, L"yes")   == 0) return true;
        if (_wcsicmp(buf, L"no")    == 0) return false;
        if (_wcsicmp(buf, L"on")    == 0) return true;
        if (_wcsicmp(buf, L"off")   == 0) return false;
        if (_wcsicmp(buf, L"1")     == 0) return true;
        if (_wcsicmp(buf, L"0")     == 0) return false;
        return dflt;
    };
    g_autoFallbackWord       = readBool(L"AutoFallback", L"Word",       true);
    g_autoFallbackExcel      = readBool(L"AutoFallback", L"Excel",      true);
    g_autoFallbackPowerPoint = readBool(L"AutoFallback", L"PowerPoint", true);

    // [PreviewHandlers] ReportPath — optional opt-in discovery aid.
    // When set to a non-empty file path the host writes a snapshot of
    // every installed preview handler and every per-extension
    // assignment to that file on each launch.  Off by default; only
    // useful when first configuring an override.
    GetPrivateProfileStringW(L"PreviewHandlers", L"ReportPath", L"",
                             buf, ARRAYSIZE(buf), iniPath.c_str());
    g_handlersReportPath = ExpandEnv(buf);

    // [PreviewHandlers] — per-extension CLSID overrides.
    //
    // Format inside the INI:
    //   [PreviewHandlers]
    //   .pdf={1531D583-8375-4D3F-B5FB-D23BBD169F22}
    //   .msg={70AE7BC8-D953-409F-9DC2-2BE5F77AF6AB}
    //
    // Keys are extensions (with or without leading dot, case-insensitive).
    // Values are full GUIDs in braces.  Trailing inline ";comment" and
    // surrounding whitespace are stripped, so this also works:
    //   .pdf = {1531D583-8375-4D3F-B5FB-D23BBD169F22}   ; Edge PDF
    //
    // Each successfully-loaded entry is logged (if logging is enabled).
    // Rejected entries (bad CLSID format, empty key, …) are logged too
    // so the user can spot typos.
    //
    // We read the entire section with GetPrivateProfileSectionW and walk
    // the returned double-null-terminated key=value pairs ourselves so
    // we can also pick up `ReportPath=…` (above) from the same section
    // without it ending up in the override map.
    auto trim = [](std::wstring& s) {
        while (!s.empty() && (s.back() == L' ' || s.back() == L'\t'))
            s.pop_back();
        size_t lead = 0;
        while (lead < s.size() && (s[lead] == L' ' || s[lead] == L'\t'))
            ++lead;
        if (lead) s.erase(0, lead);
    };

    g_extensionOverrides.clear();
    g_deniedExtensions.clear();
    {
        constexpr DWORD kSectionBufWChars = 8192;
        auto* sectionBuf = new wchar_t[kSectionBufWChars]();
        DWORD got = GetPrivateProfileSectionW(L"PreviewHandlers",
                                              sectionBuf, kSectionBufWChars,
                                              iniPath.c_str());
        if (got > 0)
        {
            for (const wchar_t* p = sectionBuf; *p; p += wcslen(p) + 1)
            {
                const wchar_t* eq = wcschr(p, L'=');
                if (!eq || eq == p) continue;

                std::wstring key(p, eq - p);
                std::wstring val(eq + 1);

                // Strip inline ';comment' from the value before trimming.
                // (Keys don't need this — '=' already terminates them.)
                auto semi = val.find(L';');
                if (semi != std::wstring::npos) val.erase(semi);

                trim(key);
                trim(val);

                // Normalise key: lowercase, ensure leading dot.
                for (auto& c : key) c = (wchar_t)towlower(c);
                if (!key.empty() && key[0] != L'.') key.insert(0, L".");

                // Skip [PreviewHandlers]-section keys that aren't
                // extension overrides (notably ReportPath, which has
                // already been read above via GetPrivateProfileStringW).
                if (_wcsicmp(key.c_str(), L".reportpath") == 0) continue;

                if (key.size() < 2)
                {
                    HostLog(L"  PreviewHandlers: skipping malformed key '%s'", key.c_str());
                    continue;
                }

                // Empty value → deny entry.  The extension is treated
                // as if no preview handler were registered for it
                // (FindPreviewHandlerClsid will return REGDB_E_CLASSNOTREG).
                if (val.empty())
                {
                    g_deniedExtensions.insert(key);
                    HostLog(L"  PreviewHandlers deny: %s (treated as no handler)",
                            key.c_str());
                    continue;
                }

                // Non-empty value: must parse as a CLSID.
                CLSID dummy;
                if (SUCCEEDED(CLSIDFromString(val.c_str(), &dummy)))
                {
                    g_extensionOverrides[key] = val;
                    HostLog(L"  PreviewHandlers override: %s -> %s",
                            key.c_str(), val.c_str());
                }
                else
                {
                    HostLog(L"  PreviewHandlers: REJECTED (bad CLSID) %s = %s",
                            key.c_str(), val.c_str());
                }
            }
        }
        delete[] sectionBuf;
    }

    return true;
}

static void LoadConfig()
{
    // 1) Per-user override under TC's standard config directory.
    //    APPDATA itself is always short (~30 chars), but using a fixed
    //    MAX_PATH buffer is the kind of inconsistency that drifts into
    //    real bugs over time; a two-call dynamic read is just as cheap.
    DWORD needed = GetEnvironmentVariableW(L"APPDATA", nullptr, 0);
    if (needed > 0)
    {
        std::wstring appdata(needed, L'\0');
        DWORD got = GetEnvironmentVariableW(L"APPDATA", appdata.data(), needed);
        if (got > 0 && got < needed)
        {
            appdata.resize(got);
            std::wstring userIni = appdata + L"\\GHISLER\\TCOfficeView.ini";
            if (LoadConfigFrom(userIni)) return;
        }
    }
    // 2) System-wide INI shipped alongside the host EXE.
    LoadConfigFrom(GetHostExeDir() + L"\\TCOfficeView.ini");
}

static void EnsureParentDir(const std::wstring& filePath)
{
    auto sep = filePath.find_last_of(L"\\/");
    if (sep == std::wstring::npos) return;
    std::wstring dir = filePath.substr(0, sep);
    SHCreateDirectoryExW(nullptr, dir.c_str(), nullptr);
}

// ---------------------------------------------------------------------------
// Long-path support
//
// Win32 file APIs that don't internally use the long-path-aware variants
// silently fail (typically with ERROR_PATH_NOT_FOUND or some "file not
// found" HRESULT) when the path exceeds ~MAX_PATH characters.  The
// canonical workaround is the \\?\ prefix, which tells the kernel to
// pass the path through without normalization and without enforcing the
// legacy 260-char limit.  The prefix has rules:
//
//   - Only legal on FULLY QUALIFIED, normalized paths (no "." or "..",
//     no forward slashes).  TC always hands us such paths, so we trust
//     the input and don't try to normalize it ourselves.
//   - Drive-letter form:  C:\foo\bar       →  \\?\C:\foo\bar
//   - UNC form:           \\server\share\… →  \\?\UNC\server\share\…
//
// We only prepend the prefix when the path actually exceeds the safe
// MAX_PATH ceiling.  Applying it to every path is harmless for the file
// APIs themselves but disables Shell-side niceties (icon overlays,
// link tracking, COM relative-path resolution), so we keep short paths
// untouched.  The threshold of MAX_PATH-12 is what Microsoft's
// own documentation recommends — it leaves room for an "8.3" filename
// suffix that some APIs append internally.
// ---------------------------------------------------------------------------

static std::wstring EnsureLongPathPrefix(const std::wstring& path)
{
    // Already prefixed — \\?\, \\?\UNC\, \\.\ all bypass the limit.
    if (path.size() >= 4 && path[0] == L'\\' && path[1] == L'\\' &&
        (path[2] == L'?' || path[2] == L'.') && path[3] == L'\\')
        return path;

    // Short enough that no Win32 API would refuse it.
    if (path.size() < MAX_PATH - 12) return path;

    // UNC path: \\server\share\… → \\?\UNC\server\share\…
    if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\')
        return L"\\\\?\\UNC\\" + path.substr(2);

    // Drive-letter / device path.
    return L"\\\\?\\" + path;
}

// ---------------------------------------------------------------------------
// Diagnostic logging
//
// At compile time, HOST_LOG_ENABLED = 0 strips the logging code entirely.
// At run time, logging is gated by g_logPath (loaded from TCOfficeView.ini):
// an empty path means "do nothing", a non-empty path means "append every
// HostLog call to that file, creating missing parent directories on first
// write." See the LoadConfig section above for the INI lookup order.
// ---------------------------------------------------------------------------

#define HOST_LOG_ENABLED 1

#if HOST_LOG_ENABLED
static HANDLE g_hLog = nullptr;
static CRITICAL_SECTION g_logCs;

static void HostLog(const wchar_t* fmt, ...)
{
    EnterCriticalSection(&g_logCs);
    if (g_logPath.empty())
    {
        LeaveCriticalSection(&g_logCs);
        return;
    }
    if (!g_hLog)
    {
        EnsureParentDir(g_logPath);
        g_hLog = CreateFileW(g_logPath.c_str(), FILE_APPEND_DATA,
                             FILE_SHARE_READ | FILE_SHARE_WRITE,
                             nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    }
    if (g_hLog && g_hLog != INVALID_HANDLE_VALUE)
    {
        wchar_t prefix[64] = {};
        SYSTEMTIME st;
        GetLocalTime(&st);
        _snwprintf_s(prefix, _TRUNCATE,
                     L"[%02d:%02d:%02d.%03d pid=%lu] ",
                     st.wHour, st.wMinute, st.wSecond, st.wMilliseconds,
                     GetCurrentProcessId());

        wchar_t body[1024] = {};
        va_list ap;
        va_start(ap, fmt);
        _vsnwprintf_s(body, _TRUNCATE, fmt, ap);
        va_end(ap);

        DWORD written = 0;
        WriteFile(g_hLog, prefix,
                  static_cast<DWORD>(wcslen(prefix) * sizeof(wchar_t)),
                  &written, nullptr);
        WriteFile(g_hLog, body,
                  static_cast<DWORD>(wcslen(body) * sizeof(wchar_t)),
                  &written, nullptr);
        static const wchar_t nl[] = L"\r\n";
        WriteFile(g_hLog, nl, sizeof(nl) - sizeof(wchar_t), &written, nullptr);
        FlushFileBuffers(g_hLog);
    }
    LeaveCriticalSection(&g_logCs);
}
#else
static void HostLog(const wchar_t*, ...) {}
#endif

// ---------------------------------------------------------------------------
// Long-path workaround via NTFS directory junctions.
//
// Despite the host being longPathAware and the system having
// LongPathsEnabled=1, several real-world consumers still refuse paths
// past MAX_PATH-1:
//
//   - Microsoft Office's Word preview handler hard-rejects long paths
//     via IInitializeWithFile::Initialize → E_NOTIMPL (looks like an
//     internal StringCchCopyW(buf, MAX_PATH, path) check).
//   - Excel and PowerPoint, both as preview handlers and at the
//     application level via Documents.Open / Workbooks.Open, refuse
//     paths longer than MAX_PATH-1 internally.
//
// The classic workaround is to create a directory junction (an NTFS
// reparse point with the IO_REPARSE_TAG_MOUNT_POINT tag) in %TEMP%
// pointing at the file's parent folder.  Both the junction path and
// "junction\filename" are well under MAX_PATH, so every consumer
// sees a short path and is happy.  The kernel transparently follows
// the junction to the real file.  We don't need elevated privileges
// (unlike symbolic links).
// ---------------------------------------------------------------------------

#ifndef IO_REPARSE_TAG_MOUNT_POINT
#define IO_REPARSE_TAG_MOUNT_POINT 0xA0000003L
#endif

// Minimal in-memory layout of REPARSE_DATA_BUFFER for mount points.
// The full union is in ntifs.h (DDK only).  We just need the
// IO_REPARSE_TAG_MOUNT_POINT variant; symbolic links and the generic
// buffer are deliberately omitted.
struct JunctionReparseBuffer
{
    ULONG  ReparseTag;
    USHORT ReparseDataLength;
    USHORT Reserved;
    USHORT SubstituteNameOffset;
    USHORT SubstituteNameLength;
    USHORT PrintNameOffset;
    USHORT PrintNameLength;
    WCHAR  PathBuffer[1];
};

// Create a directory junction at `junctionPath` pointing to `targetPath`.
// `junctionPath` must not yet exist (we create the directory).
// `targetPath` must be an existing directory.  Logs and falls back
// gracefully on any step's failure.
static bool CreateJunctionSta(const std::wstring& junctionPath,
                              const std::wstring& targetPath)
{
    if (!CreateDirectoryW(junctionPath.c_str(), nullptr))
    {
        HostLog(L"  CreateJunction: CreateDirectoryW('%s') failed err=%lu",
                junctionPath.c_str(), GetLastError());
        return false;
    }

    HANDLE h = CreateFileW(
        junctionPath.c_str(),
        GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE,
        nullptr,
        OPEN_EXISTING,
        FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT,
        nullptr);
    if (h == INVALID_HANDLE_VALUE)
    {
        HostLog(L"  CreateJunction: CreateFileW('%s') failed err=%lu",
                junctionPath.c_str(), GetLastError());
        RemoveDirectoryW(junctionPath.c_str());
        return false;
    }

    // Reparse point payload:
    //   SubstituteName — NT-namespace path "\??\<target>", used by the
    //                    kernel to resolve the junction.
    //   PrintName      — display path "<target>", shown to user tools.
    // Both null-terminated, laid out back-to-back in PathBuffer.
    const std::wstring substName = L"\\??\\" + targetPath;
    const std::wstring& printName = targetPath;

    const USHORT substBytes = static_cast<USHORT>(substName.size() * sizeof(wchar_t));
    const USHORT printBytes = static_cast<USHORT>(printName.size() * sizeof(wchar_t));

    const SIZE_T pathBufferBytes =
        substBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);
    const SIZE_T totalSize =
        offsetof(JunctionReparseBuffer, PathBuffer) + pathBufferBytes;

    std::vector<BYTE> bytes(totalSize, 0);
    auto* rb = reinterpret_cast<JunctionReparseBuffer*>(bytes.data());
    rb->ReparseTag           = IO_REPARSE_TAG_MOUNT_POINT;
    // ReparseDataLength counts everything AFTER the 8-byte common header
    // (ReparseTag + ReparseDataLength + Reserved).
    rb->ReparseDataLength    = static_cast<USHORT>(totalSize - 8);
    rb->Reserved             = 0;
    rb->SubstituteNameOffset = 0;
    rb->SubstituteNameLength = substBytes;
    rb->PrintNameOffset      = static_cast<USHORT>(substBytes + sizeof(wchar_t));
    rb->PrintNameLength      = printBytes;
    memcpy(rb->PathBuffer, substName.c_str(), substBytes + sizeof(wchar_t));
    memcpy(reinterpret_cast<BYTE*>(rb->PathBuffer) + substBytes + sizeof(wchar_t),
           printName.c_str(), printBytes + sizeof(wchar_t));

    DWORD bytesReturned = 0;
    BOOL ok = DeviceIoControl(h, FSCTL_SET_REPARSE_POINT,
                              rb, static_cast<DWORD>(totalSize),
                              nullptr, 0, &bytesReturned, nullptr);
    DWORD err = ok ? 0 : GetLastError();
    CloseHandle(h);

    if (!ok)
    {
        HostLog(L"  CreateJunction: DeviceIoControl failed err=%lu", err);
        RemoveDirectoryW(junctionPath.c_str());
        return false;
    }

    HostLog(L"  CreateJunction OK: '%s' -> '%s'",
            junctionPath.c_str(), targetPath.c_str());
    return true;
}

// Category GUID for Windows Preview Handlers — the subkey
// HKCR\.<ext>\shellex\{8895b1c6-...} points to the handler's CLSID.
static const GUID kPreviewHandlerCategory =
    { 0x8895b1c6, 0xb41f, 0x4c1c, { 0xa5, 0x62, 0x0d, 0x56, 0x42, 0x50, 0x83, 0x6f } };

// ---------------------------------------------------------------------------
// Application classification (must be declared before HostState).
// ---------------------------------------------------------------------------

enum class AppKind { Other, Word, Excel, PowerPoint };

// ---------------------------------------------------------------------------
// Process-wide state. Modified only from the STA thread.
// ---------------------------------------------------------------------------

struct HostState
{
    HWND                hwndPluginChild = nullptr;        // window inside TC, supplied on cmd line
    HWND                hwndRender      = nullptr;        // our own child, reparented into hwndPluginChild
    HWND                hwndSta         = nullptr;        // message-only window for cross-thread dispatch
    HWND                hwndFallback    = nullptr;        // read-only EDIT shown when no preview handler is usable
    HFONT               hFallbackFont   = nullptr;        // font owned by us, freed when hwndFallback is hidden
    HWND                hwndUnblockButton = nullptr;      // "Unblock & retry" button shown when the file has MOTW
    HFONT               hUnblockButtonFont = nullptr;     // font for the unblock button
    HANDLE              hPipe           = INVALID_HANDLE_VALUE;
    IUnknown*           pHandlerUnk     = nullptr;        // raw COM object
    IPreviewHandler*    pHandler        = nullptr;        // queried IPreviewHandler view

    // Full-mode (OLE Automation) state, one slot per Office app. At most one
    // of (Word, Excel, PowerPoint) is live at a time — switching to a file
    // of a different type quits the previously loaded app. Within a Lister
    // session, switches between files of the *same* app reuse the running
    // Application instance and just close/open the document.
    IDispatch*          pWordApp        = nullptr;
    IDispatch*          pWordDoc        = nullptr;
    HWND                hwndWordApp     = nullptr;        // Word main window reparented into hwndRender
    IDispatch*          pExcelApp       = nullptr;
    IDispatch*          pExcelWb        = nullptr;
    HWND                hwndExcelApp    = nullptr;        // Excel main window reparented into hwndRender
    // Mode-switch button (always-visible Win32 BUTTON in the top-right of
    // the render pane). It lets the user temporarily flip the current
    // preview between quick and full mode without changing the INI default.
    // The switch is per-preview only — loading a different file or
    // re-opening this one starts over at the INI-configured mode.
    HWND                hwndModeButton  = nullptr;
    HFONT               hModeButtonFont = nullptr;
    std::wstring        currentFile;                    // last LOAD argument, for re-load on switch
    AppKind             currentFileApp  = AppKind::Other;
    Mode                currentLoadedMode = Mode::Quick;

    // Guard against re-entrant LOAD/SWITCH_MODE while a previous load is
    // still running. The retry loops inside LoadXxxFullSta pump messages
    // (PeekMessage) so a second button click could otherwise recurse into
    // LoadFileWithModeSta and corrupt COM state.
    bool                loadingInProgress = false;

    // Invisible overlay window that sits in the top-right corner of the
    // render pane and swallows mouse clicks aimed at the Office app's own
    // close button. Office windows run out-of-process, so SetWindowSubclass
    // does not work on them; this guard window (a child of hwndRender in
    // our own process) is the only reliable way to block the button.
    HWND                hwndCloseGuard  = nullptr;

    // Job Object that owns the Office processes we spawn via COM. With
    // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE set, the kernel kills every
    // process in the job the moment its last handle (ours, here)
    // closes — so however this host process dies (clean WM_HOST_CLOSE,
    // TerminateProcess from the plugin DLL, Windows shutdown), the
    // Excel / Word / PowerPoint instances we created go with it
    // instead of being left orphaned.
    HANDLE              hOfficeJob      = nullptr;

    // Long-path workaround.  Most Office consumers (Word preview
    // handler, Excel & PowerPoint at the application level) bail on
    // paths longer than MAX_PATH-1 even with longPathAware + the
    // system-wide LongPathsEnabled registry switch.  When we see such
    // a path we create a directory junction in %TEMP% pointing at the
    // file's parent and pass the resulting short alias path down to
    // every consumer.  This field holds the junction's own path so
    // we can RemoveDirectoryW it on the next LOAD / on CLOSE.
    std::wstring              activeJunctionDir;

    // Junctions whose RemoveDirectoryW failed (typically because Office
    // still has a monitoring / change-notification handle on the
    // directory).  Retried at every cleanup opportunity; whatever
    // remains at WM_HOST_CLOSE is logged and left for the next host
    // launch's startup sweep to pick up.
    std::vector<std::wstring> staleJunctionDirs;

    IDispatch*          pPptApp         = nullptr;
    IDispatch*          pPptPres        = nullptr;
    HWND                hwndPptApp      = nullptr;        // PowerPoint main window reparented into hwndRender
    // True when CoCreateInstance(PowerPoint.Application) connected to a
    // pre-existing PPT process rather than creating a new one (MULTIPLEUSE
    // registration).  In that case we must NOT call Application.Quit (which
    // would close all the user's open presentations) and must NOT assign the
    // process to our kill-on-close job object.
    bool                pPptAppIsShared = false;
};

static HostState g_state;

// Message IDs the worker thread posts to the STA window.
//
// We use PostMessage (not SendMessage) because cross-thread SendMessage puts
// the receiving WndProc into the "input-synchronous" state, and COM refuses
// outgoing calls from there with RPC_E_CANTCALLOUT_ININPUTSYNCCALL. With
// PostMessage the WndProc runs in a normal message-pump state and
// CoCreateInstance works.
//
// Ownership of lParam payloads (for WM_HOST_LOAD) transfers from the worker
// to the WndProc handler, which is responsible for freeing.
static constexpr UINT WM_HOST_LOAD   = WM_USER + 1;     // lParam = wchar_t* (heap, delete[])
static constexpr UINT WM_HOST_RESIZE = WM_USER + 2;     // wParam = MAKELONG(w, h)
static constexpr UINT WM_HOST_CLOSE       = WM_USER + 3;
static constexpr UINT WM_HOST_SWITCH_MODE = WM_USER + 4;     // user clicked the mode-switch button
static constexpr UINT WM_HOST_UNBLOCK_AND_RELOAD = WM_USER + 5;  // user clicked the MOTW Unblock button

// Control ID for the mode-switch BUTTON child window of hwndRender —
// referenced from RenderWndProc's WM_COMMAND handler.
static constexpr UINT_PTR kModeButtonId           = 1001;
// One-shot timer: re-raise the mode button after a resize. Preview
// handlers and Office apps may asynchronously reposition their windows
// in response to SetRect / SetWindowPos, burying the button.
static constexpr UINT_PTR kModeButtonZTimerId     = 1002;
// Periodic timer (~100 ms): keeps the button above sibling windows for the
// entire life of the render window.  Needed because preview handlers (Word
// in particular) can call SetWindowPos on their own window — e.g. while
// scrolling — without SWP_NOZORDER, which moves them above the button
// without generating WM_SIZE on hwndRender.
static constexpr UINT_PTR kModeButtonKeepTopTimerId = 1003;

// Mark-of-the-Web fallback panel: optional "Unblock & retry" button shown
// alongside the read-only EDIT when the file the handler refused to
// render is marked as having been downloaded from the internet.  Click
// deletes the file's Zone.Identifier alternate data stream and reloads.
static constexpr UINT_PTR kUnblockButtonId          = 1004;

// ---------------------------------------------------------------------------
// Registry lookup
// ---------------------------------------------------------------------------

static std::wstring GuidToBraces(REFGUID g)
{
    wchar_t buf[64] = {};
    StringFromGUID2(g, buf, ARRAYSIZE(buf));
    return buf;
}

static bool ReadDefaultString(HKEY root, LPCWSTR subPath, std::wstring& out)
{
    HKEY hKey = nullptr;
    if (RegOpenKeyExW(root, subPath, 0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return false;
    wchar_t buf[256] = {};
    DWORD cb = sizeof(buf);
    DWORD type = 0;
    LONG rc = RegQueryValueExW(hKey, nullptr, nullptr, &type,
                               reinterpret_cast<BYTE*>(buf), &cb);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS) return false;
    if (type != REG_SZ && type != REG_EXPAND_SZ) return false;
    out = buf;
    return true;
}

// System-only lookup chain — no INI override.  Used both by the public
// FindPreviewHandlerClsid wrapper (as the fallback when no override matches)
// and by the available-handlers report generator (which wants to show what
// the system would do without any of our overrides in the picture).
static HRESULT FindSystemPreviewHandlerClsid(LPCWSTR path, CLSID* out)
{
    LPCWSTR dot = wcsrchr(path, L'.');
    if (!dot || !dot[1]) return E_INVALIDARG;

    const std::wstring ext     = dot;          // e.g. ".msg"
    const std::wstring catGuid = GuidToBraces(kPreviewHandlerCategory);
    std::wstring       clsidStr;

    // Step 1 — shellex key directly on the extension.
    {
        const std::wstring sub = ext + L"\\shellex\\" + catGuid;
        if (ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
        {
            HostLog(L"FindPreviewHandlerClsid: '%s' found via direct shellex", ext.c_str());
            return CLSIDFromString(clsidStr.c_str(), out);
        }
    }

    // Step 2 — shellex on the extension's default ProgID.
    {
        std::wstring progId;
        if (ReadDefaultString(HKEY_CLASSES_ROOT, ext.c_str(), progId) && !progId.empty())
        {
            const std::wstring sub = progId + L"\\shellex\\" + catGuid;
            if (ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
            {
                HostLog(L"FindPreviewHandlerClsid: '%s' found via ProgID '%s'",
                        ext.c_str(), progId.c_str());
                return CLSIDFromString(clsidStr.c_str(), out);
            }
        }
    }

    // Step 3 — shellex on each ProgID listed under OpenWithProgids.
    //
    // Outlook registers its MSG preview handler only under one of the
    // versioned ProgIDs in this subkey (e.g. Outlook.File.msg.16), not as
    // the default ProgID of ".msg" — so Steps 1 and 2 miss it.
    {
        const std::wstring owpPath = ext + L"\\OpenWithProgids";
        HKEY hOwp = nullptr;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, owpPath.c_str(), 0, KEY_READ, &hOwp)
                == ERROR_SUCCESS)
        {
            wchar_t name[512] = {};
            DWORD   nameLen   = ARRAYSIZE(name);
            for (DWORD idx = 0;
                 RegEnumValueW(hOwp, idx, name, &nameLen,
                               nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS;
                 ++idx, nameLen = ARRAYSIZE(name))
            {
                if (!name[0]) continue;
                const std::wstring sub = std::wstring(name) + L"\\shellex\\" + catGuid;
                if (ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
                {
                    RegCloseKey(hOwp);
                    HostLog(L"FindPreviewHandlerClsid: '%s' found via OpenWithProgids '%s'",
                            ext.c_str(), name);
                    return CLSIDFromString(clsidStr.c_str(), out);
                }
            }
            RegCloseKey(hOwp);
        }
    }

    // Step 4 — system-level association on the extension itself.
    //
    // HKCR\SystemFileAssociations\<ext> is where Windows places system-wide
    // shell extensions that should apply to every program registered for a
    // given extension.  Some installs (e.g. modern Outlook for MSG) only
    // register the preview handler here.
    {
        const std::wstring sub = L"SystemFileAssociations\\" + ext
                               + L"\\shellex\\" + catGuid;
        if (ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
        {
            HostLog(L"FindPreviewHandlerClsid: '%s' found via SystemFileAssociations\\<ext>",
                    ext.c_str());
            return CLSIDFromString(clsidStr.c_str(), out);
        }
    }

    // Step 5 — system-level association on the extension's PerceivedType.
    //
    // HKCR\<ext>\PerceivedType groups extensions into broad categories
    // (e.g. "document", "image", "audio").  Handlers registered under
    // HKCR\SystemFileAssociations\<perceived-type> apply to every extension
    // tagged with that type.
    {
        std::wstring perceivedType;
        if (ReadDefaultString(HKEY_CLASSES_ROOT,
                              (ext + L"\\PerceivedType").c_str(),
                              perceivedType) && !perceivedType.empty())
        {
            const std::wstring sub = L"SystemFileAssociations\\" + perceivedType
                                   + L"\\shellex\\" + catGuid;
            if (ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
            {
                HostLog(L"FindPreviewHandlerClsid: '%s' found via PerceivedType '%s'",
                        ext.c_str(), perceivedType.c_str());
                return CLSIDFromString(clsidStr.c_str(), out);
            }
        }
    }

    HostLog(L"FindPreviewHandlerClsid: no preview handler found for '%s'", ext.c_str());
    return REGDB_E_CLASSNOTREG;
}

// Public lookup: tries the INI [PreviewHandlers] override first, then
// falls back to the system-wide registry chain.  Used for actual file
// previews — wherever we'd resolve "what handler do I instantiate for
// this file" the override gets a chance to redirect us.
static HRESULT FindPreviewHandlerClsid(LPCWSTR path, CLSID* out)
{
    LPCWSTR dot = wcsrchr(path, L'.');
    if (!dot || !dot[1]) return E_INVALIDARG;

    // Lowercased extension is used for both the deny set and the
    // override map lookups below.
    std::wstring ext = dot;
    for (auto& c : ext) c = (wchar_t)towlower(c);

    // Deny list: ".ext=" (empty value) in [PreviewHandlers] makes the
    // extension behave as if no handler were registered.  The plugin
    // DLL applies the same check before forwarding LOAD to the host,
    // so we should rarely reach this branch — it's a defence-in-depth
    // backstop in case the DLL's deny list is somehow stale
    // (e.g. INI edited between the DLL's read and the host spawn).
    if (!g_deniedExtensions.empty() && g_deniedExtensions.count(ext))
    {
        HostLog(L"FindPreviewHandlerClsid: '%s' DENIED via INI [PreviewHandlers]",
                ext.c_str());
        return REGDB_E_CLASSNOTREG;
    }

    // INI override.  Lets the user pick a specific preview handler for
    // an extension regardless of what's registered system-wide (typical
    // use: Adobe Reader claimed .pdf but the user prefers Edge's PDF
    // handler for our preview pane).  Bad CLSIDs were already filtered
    // out at config load, so a parse failure here can only mean a CLSID
    // string that was valid syntactically but isn't actually a
    // registered COM class — we still try it and let CoCreateInstance
    // complain if so.
    if (!g_extensionOverrides.empty())
    {
        auto it = g_extensionOverrides.find(ext);
        if (it != g_extensionOverrides.end())
        {
            HRESULT hr = CLSIDFromString(it->second.c_str(), out);
            HostLog(L"FindPreviewHandlerClsid: '%s' resolved via INI override "
                    L"PreviewHandlers=%s (0x%08lX)",
                    ext.c_str(), it->second.c_str(), static_cast<long>(hr));
            if (SUCCEEDED(hr)) return S_OK;
            // Fall through to the registry-based chain on parse failure.
        }
    }

    return FindSystemPreviewHandlerClsid(path, out);
}

// ---------------------------------------------------------------------------
// Available-handlers report.
//
// Generates a human-readable, copy-pasteable file at
// %APPDATA%\GHISLER\TCOfficeView.available-handlers.txt every time the
// host starts.  Two sections:
//
//   1) Every preview handler installed on this machine (the master
//      list users pick CLSIDs from), enumerated from
//      HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers.
//   2) For every file extension that has *any* preview handler today,
//      the CLSID currently in effect — preformatted as commented-out
//      INI lines so users can copy a row into
//      [PreviewHandlers] in TCOfficeView.ini, uncomment it, and
//      change the CLSID to one of the alternatives from section 1.
//
// Runs on a background thread so it doesn't add to F3 cold-start
// latency.  Atomic write via .tmp + MoveFileExW.
// ---------------------------------------------------------------------------

// Enumerate HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers
// into a map keyed by uppercase CLSID-with-braces.  Values are friendly
// names ("Microsoft Edge PDF Preview Handler", …) as registered there.
static std::map<std::wstring, std::wstring> EnumerateInstalledPreviewHandlersFromHklm()
{
    std::map<std::wstring, std::wstring> result;
    HKEY hKey = nullptr;
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE,
                      L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers",
                      0, KEY_READ, &hKey) != ERROR_SUCCESS)
        return result;

    for (DWORD idx = 0;; ++idx)
    {
        wchar_t name[256] = {};
        wchar_t data[1024] = {};
        DWORD nameLen = ARRAYSIZE(name);
        DWORD dataBytes = sizeof(data) - sizeof(wchar_t);   // leave room for terminator
        DWORD type = 0;
        LONG rc = RegEnumValueW(hKey, idx, name, &nameLen, nullptr,
                                &type, reinterpret_cast<BYTE*>(data), &dataBytes);
        if (rc == ERROR_NO_MORE_ITEMS) break;
        if (rc != ERROR_SUCCESS) continue;
        if (type != REG_SZ && type != REG_EXPAND_SZ) continue;

        std::wstring clsid(name);
        // Normalise CLSID to uppercase for stable map keys.
        for (auto& c : clsid) c = (wchar_t)towupper(c);
        result[clsid] = std::wstring(data);
    }
    RegCloseKey(hKey);
    return result;
}

// Walk every "*.ext" key under HKEY_CLASSES_ROOT and, for each one that
// resolves to a registered preview handler via the standard system
// lookup chain (no INI overrides), record extension → CLSID.
//
// Keys are normalised to lowercase ".ext" form; CLSIDs are uppercase
// {…} form so they cross-reference cleanly with the HKLM map above.
static std::map<std::wstring, std::wstring> EnumerateSystemHandlerAssignments()
{
    std::map<std::wstring, std::wstring> result;
    HKEY hKey = nullptr;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, L"", 0,
                      KEY_READ | KEY_ENUMERATE_SUB_KEYS, &hKey) != ERROR_SUCCESS)
        return result;

    for (DWORD idx = 0;; ++idx)
    {
        wchar_t subKey[256] = {};
        DWORD subKeyLen = ARRAYSIZE(subKey);
        LONG rc = RegEnumKeyExW(hKey, idx, subKey, &subKeyLen,
                                nullptr, nullptr, nullptr, nullptr);
        if (rc == ERROR_NO_MORE_ITEMS) break;
        if (rc != ERROR_SUCCESS) continue;
        if (subKey[0] != L'.') continue;                    // not an extension entry

        // Synthesise a fake path so we can reuse FindSystemPreviewHandlerClsid.
        std::wstring fakePath = L"x";
        fakePath += subKey;

        CLSID clsid;
        if (FindSystemPreviewHandlerClsid(fakePath.c_str(), &clsid) == S_OK)
        {
            wchar_t clsidStr[64] = {};
            if (StringFromGUID2(clsid, clsidStr, ARRAYSIZE(clsidStr)) > 0)
            {
                std::wstring ext = subKey;
                for (auto& c : ext)     c = (wchar_t)towlower(c);    // ".PDF" → ".pdf"
                std::wstring clsidUpper(clsidStr);
                for (auto& c : clsidUpper) c = (wchar_t)towupper(c);
                result[ext] = clsidUpper;
            }
        }
    }
    RegCloseKey(hKey);
    return result;
}

// Append `s` to `outUtf8` as UTF-8 bytes.
static void AppendWideAsUtf8(std::string& outUtf8, const std::wstring& s)
{
    if (s.empty()) return;
    int n = WideCharToMultiByte(CP_UTF8, 0, s.c_str(),
                                static_cast<int>(s.size()),
                                nullptr, 0, nullptr, nullptr);
    if (n <= 0) return;
    size_t start = outUtf8.size();
    outUtf8.resize(start + n);
    WideCharToMultiByte(CP_UTF8, 0, s.c_str(), static_cast<int>(s.size()),
                        outUtf8.data() + start, n, nullptr, nullptr);
}

static void AppendLineUtf8(std::string& outUtf8, const std::wstring& s)
{
    AppendWideAsUtf8(outUtf8, s);
    outUtf8 += "\r\n";
}

// Compose and write %APPDATA%\GHISLER\TCOfficeView.available-handlers.txt.
// Best-effort: any failure (missing APPDATA, write error, file locked by
// an editor) is logged and ignored — the report is a convenience, not a
// requirement for the plugin to work.
static void WriteAvailableHandlersReportSta()
{
    // Opt-in via INI: [PreviewHandlers] ReportPath=<path>.  Empty
    // path means the user hasn't asked for the report, so we don't
    // generate anything (cheap silent skip on every host launch).
    if (g_handlersReportPath.empty()) return;

    const std::wstring& destPath = g_handlersReportPath;
    const std::wstring  tmpPath  = destPath + L".tmp";

    // Create any missing parent directories so the user's chosen path
    // can be a fresh location (e.g. %LocalAppData%\TCOfficeView\…).
    EnsureParentDir(destPath);

    auto installed   = EnumerateInstalledPreviewHandlersFromHklm();
    auto assignments = EnumerateSystemHandlerAssignments();

    HostLog(L"AvailableHandlers report: %zu installed, %zu extension assignments",
            installed.size(), assignments.size());

    // Build the report content in memory as UTF-8.
    std::string body;
    body.reserve(8192);
    // UTF-8 BOM so Notepad and other editors detect encoding correctly.
    body.append("\xEF\xBB\xBF");

    SYSTEMTIME st;
    GetLocalTime(&st);
    wchar_t header[256] = {};
    _snwprintf_s(header, _TRUNCATE,
                 L"TCOfficeView — installed preview handlers and current per-extension assignments\r\n"
                 L"Regenerated automatically every time the plugin host starts.\r\n"
                 L"Last update: %04u-%02u-%02u %02u:%02u:%02u local time.\r\n",
                 st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    AppendWideAsUtf8(body, header);

    AppendLineUtf8(body, L"");
    AppendLineUtf8(body, L"================================================================");
    AppendLineUtf8(body, L"1) Installed preview handlers on this machine");
    AppendLineUtf8(body, L"================================================================");
    AppendLineUtf8(body, L"   Source: HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers");
    AppendLineUtf8(body, L"   Pick a CLSID below to override an extension's default handler");
    AppendLineUtf8(body, L"   (use the value from the second column inside [PreviewHandlers]).");
    AppendLineUtf8(body, L"");

    if (installed.empty())
    {
        AppendLineUtf8(body, L"   (no preview handlers are registered on this machine)");
    }
    else
    {
        // Sort by friendly name for predictable, scannable output.
        std::vector<std::pair<std::wstring, std::wstring>> byName;     // (friendly, clsid)
        byName.reserve(installed.size());
        for (const auto& kv : installed) byName.emplace_back(kv.second, kv.first);
        std::sort(byName.begin(), byName.end(),
                  [](const auto& a, const auto& b) {
                      return _wcsicmp(a.first.c_str(), b.first.c_str()) < 0;
                  });

        size_t nameColWidth = 0;
        for (const auto& p : byName)
            if (p.first.size() > nameColWidth) nameColWidth = p.first.size();
        if (nameColWidth > 60) nameColWidth = 60;       // keep lines readable

        for (const auto& p : byName)
        {
            std::wstring line = L"   " + p.first;
            while (line.size() < nameColWidth + 6) line += L' ';
            line += p.second;
            AppendLineUtf8(body, line);
        }
    }

    AppendLineUtf8(body, L"");
    AppendLineUtf8(body, L"================================================================");
    AppendLineUtf8(body, L"2) Current per-extension assignments (system-wide defaults)");
    AppendLineUtf8(body, L"================================================================");
    AppendLineUtf8(body, L"   These are the handlers TCOfficeView picks today for each file");
    AppendLineUtf8(body, L"   type that has any preview handler registered, before any of");
    AppendLineUtf8(body, L"   your INI overrides apply.  Each line is preformatted as a");
    AppendLineUtf8(body, L"   commented-out INI entry — copy any line into");
    AppendLineUtf8(body, L"   [PreviewHandlers] in %APPDATA%\\GHISLER\\TCOfficeView.ini,");
    AppendLineUtf8(body, L"   remove the leading ';', then change the CLSID to one of the");
    AppendLineUtf8(body, L"   handlers listed in section 1 above.");
    AppendLineUtf8(body, L"");

    if (assignments.empty())
    {
        AppendLineUtf8(body, L"   (no extensions have a registered preview handler)");
    }
    else
    {
        // Column width for the ".ext={CLSID}" part.
        size_t lhsWidth = 0;
        for (const auto& kv : assignments)
        {
            size_t w = kv.first.size() + 1 + kv.second.size() + 1;   // ext + "=" + clsid + ";"
            if (w > lhsWidth) lhsWidth = w;
        }

        for (const auto& kv : assignments)
        {
            // Look up friendly name (uppercase key in installed map).
            std::wstring friendlyName;
            auto fnd = installed.find(kv.second);
            if (fnd != installed.end()) friendlyName = fnd->second;

            std::wstring line = L";" + kv.first + L"=" + kv.second;
            while (line.size() < lhsWidth + 3) line += L' ';
            if (!friendlyName.empty())
            {
                line += L" ; ";
                line += friendlyName;
            }
            AppendLineUtf8(body, line);
        }
    }

    // Atomic write: dump to .tmp, then MoveFileEx with REPLACE_EXISTING.
    HANDLE h = CreateFileW(tmpPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                           nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE)
    {
        HostLog(L"AvailableHandlers report: CreateFile('%s') failed err=%lu",
                tmpPath.c_str(), GetLastError());
        return;
    }
    DWORD written = 0;
    BOOL writeOk = WriteFile(h, body.data(), static_cast<DWORD>(body.size()),
                             &written, nullptr);
    CloseHandle(h);
    if (!writeOk || written != body.size())
    {
        HostLog(L"AvailableHandlers report: WriteFile failed err=%lu written=%lu/%zu",
                GetLastError(), written, body.size());
        DeleteFileW(tmpPath.c_str());
        return;
    }
    if (!MoveFileExW(tmpPath.c_str(), destPath.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
    {
        HostLog(L"AvailableHandlers report: MoveFileEx failed err=%lu", GetLastError());
        DeleteFileW(tmpPath.c_str());
        return;
    }
    HostLog(L"AvailableHandlers report: wrote '%s'", destPath.c_str());
}

static DWORD WINAPI AvailableHandlersReportThread(LPVOID)
{
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_BELOW_NORMAL);
    WriteAvailableHandlersReportSta();
    return 0;
}

// ---------------------------------------------------------------------------
// Render window (in this process) reparented into the plugin's HWND.
// ---------------------------------------------------------------------------

static const wchar_t* kRenderClassName = L"TCOfficeViewHostRender";

// Forward-declared so RenderWndProc's WM_TIMER handler can call it.
static void UpdateModeButtonSta();

static LRESULT CALLBACK RenderWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_ERASEBKGND:
        {
            // When the fallback panel is visible we must paint the background
            // ourselves — the EDIT does not cover the full client area (there
            // is a gap reserved for the Unblock button) and returning 1 here
            // would leave that gap unpainted (black).
            if (g_state.hwndFallback && IsWindowVisible(g_state.hwndFallback))
            {
                RECT rc;
                GetClientRect(hWnd, &rc);
                FillRect(reinterpret_cast<HDC>(wp), &rc,
                         reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
                return 1;
            }
            return 1;                          // preview handler paints
        }

        case WM_COMMAND:
            // Click on one of our overlay buttons.  Bounce it to the STA
            // window so the action (which may involve COM, file I/O, …)
            // runs on the apartment that owns those resources.
            if (HIWORD(wp) == BN_CLICKED)
            {
                if (LOWORD(wp) == kModeButtonId)
                {
                    PostMessageW(g_state.hwndSta, WM_HOST_SWITCH_MODE, 0, 0);
                    return 0;
                }
                if (LOWORD(wp) == kUnblockButtonId)
                {
                    PostMessageW(g_state.hwndSta, WM_HOST_UNBLOCK_AND_RELOAD, 0, 0);
                    return 0;
                }
            }
            break;

        case WM_DRAWITEM:
        {
            // Custom paint for the BS_OWNERDRAW unblock button — themed
            // push buttons can't be coloured via WM_CTLCOLORBTN, and the
            // default grey makes the action easy to miss inside the
            // already-grey fallback panel.  Dark red signals "this is the
            // thing to click" without resorting to system warning icons.
            DRAWITEMSTRUCT* di = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (di && di->CtlType == ODT_BUTTON &&
                di->CtlID == kUnblockButtonId)
            {
                const bool pressed = (di->itemState & ODS_SELECTED) != 0;
                const bool focused = (di->itemState & ODS_FOCUS)    != 0;

                // Slightly darker fill while pressed for tactile feedback.
                COLORREF bgColor     = pressed ? RGB(0x70, 0x10, 0x10)
                                               : RGB(0xA0, 0x20, 0x20);
                COLORREF borderColor = RGB(0x50, 0x08, 0x08);
                COLORREF textColor   = RGB(0xFF, 0xFF, 0xFF);

                HBRUSH hBg = CreateSolidBrush(bgColor);
                FillRect(di->hDC, &di->rcItem, hBg);
                DeleteObject(hBg);

                HBRUSH hBorder = CreateSolidBrush(borderColor);
                FrameRect(di->hDC, &di->rcItem, hBorder);
                DeleteObject(hBorder);

                HFONT hFont = reinterpret_cast<HFONT>(
                    SendMessageW(di->hwndItem, WM_GETFONT, 0, 0));
                HGDIOBJ hOldFont = hFont ? SelectObject(di->hDC, hFont)
                                         : nullptr;
                SetBkMode(di->hDC, TRANSPARENT);
                SetTextColor(di->hDC, textColor);

                wchar_t label[256] = {};
                GetWindowTextW(di->hwndItem, label, ARRAYSIZE(label));
                DrawTextW(di->hDC, label, -1, &di->rcItem,
                          DT_CENTER | DT_VCENTER | DT_SINGLELINE);

                if (hOldFont) SelectObject(di->hDC, hOldFont);

                if (focused)
                {
                    RECT rcFocus = di->rcItem;
                    InflateRect(&rcFocus, -3, -3);
                    DrawFocusRect(di->hDC, &rcFocus);
                }
                return TRUE;
            }
            break;
        }

        case WM_SIZE:
            // A preview handler or embedded Office app may asynchronously
            // reposition its window after a resize (via a posted internal
            // message), pushing the mode button behind it.  Schedule a
            // deferred re-raise so the button ends up on top once the
            // handler has settled.
            SetTimer(hWnd, kModeButtonZTimerId, 150, nullptr);
            break;

        case WM_TIMER:
            if (wp == kModeButtonZTimerId)
            {
                KillTimer(hWnd, kModeButtonZTimerId);   // one-shot
                UpdateModeButtonSta();
            }
            else if (wp == kModeButtonKeepTopTimerId)
            {
                // Periodic Z-order keep-alive — but only do real work when
                // some sibling has actually moved above the mode button.
                // Calling SetWindowPos(HWND_TOP) unconditionally every
                // 100 ms generates cross-process WM_WINDOWPOSCHANGED
                // messages to embedded Office windows, which routes
                // through the system input thread and makes Total
                // Commander's own modal dialogs (file-overwrite confirm,
                // copy/delete progress, …) feel sluggish while full mode
                // is active.  GetWindow(GW_HWNDPREV) == NULL means we are
                // already at the top — the steady-state case, so the tick
                // is essentially free.
                if (g_state.hwndModeButton &&
                    IsWindowVisible(g_state.hwndModeButton) &&
                    GetWindow(g_state.hwndModeButton, GW_HWNDPREV) != nullptr)
                {
                    UpdateModeButtonSta();
                }
            }
            break;
    }
    return DefWindowProcW(hWnd, msg, wp, lp);
}

static void EnsureRenderClass(HINSTANCE hInst)
{
    static bool registered = false;
    if (registered) return;
    WNDCLASSW wc = {};
    wc.lpfnWndProc   = RenderWndProc;
    wc.hInstance     = hInst;
    wc.lpszClassName = kRenderClassName;
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    RegisterClassW(&wc);
    registered = true;
}


// ---------------------------------------------------------------------------
// Mode-switch overlay button.
//
// A small Win32 BUTTON sits in the top-right of the render pane and toggles
// the current preview between quick and full mode. The button is created
// once with the render window and stays around for the host's lifetime;
// its label and visibility are refreshed by UpdateModeButtonSta whenever
// a new file is loaded or the pane is resized.
//
// The switch is intentionally non-sticky: the next LOAD on a different
// file (or the same file in a fresh Lister session) goes back to the INI-
// configured default. See LoadFileSta vs LoadFileWithModeSta.
// ---------------------------------------------------------------------------

static int ScaleForDpi(int sizeAt96Dpi, UINT dpi)
{
    if (dpi == 0) dpi = USER_DEFAULT_SCREEN_DPI;
    return MulDiv(sizeAt96Dpi, dpi, USER_DEFAULT_SCREEN_DPI);
}

static bool CreateModeButtonSta()
{
    if (!g_state.hwndRender || g_state.hwndModeButton) return true;

    UINT dpi = GetDpiForWindow(g_state.hwndRender);
    int w = ScaleForDpi(64, dpi);
    int h = ScaleForDpi(24, dpi);

    g_state.hwndModeButton = CreateWindowExW(
        0, L"BUTTON", L"",
        WS_CHILD | WS_CLIPSIBLINGS | BS_PUSHBUTTON | BS_CENTER | BS_VCENTER,
        0, 0, w, h,
        g_state.hwndRender,
        reinterpret_cast<HMENU>(kModeButtonId),
        GetModuleHandleW(nullptr), nullptr);
    if (!g_state.hwndModeButton)
    {
        HostLog(L"CreateModeButtonSta: CreateWindowEx failed err=%lu",
                GetLastError());
        return false;
    }

    int fontHeight = -ScaleForDpi(9, dpi);       // ~9pt at the window's DPI
    g_state.hModeButtonFont = CreateFontW(
        fontHeight, 0, 0, 0,
        FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    if (g_state.hModeButtonFont)
        SendMessageW(g_state.hwndModeButton, WM_SETFONT,
                     reinterpret_cast<WPARAM>(g_state.hModeButtonFont), TRUE);

    // Stays hidden until UpdateModeButtonSta gets called from the first LOAD.
    return true;
}

// Forward-declared here because UpdateModeButtonSta calls it and
// SelectMode's actual definition appears later in the file.
static Mode SelectMode(AppKind app);

// Forward-declared so BuildFallbackText can tailor its hint for Office
// files (Word / Excel / PowerPoint) without taking the AppKind as a
// parameter.  Real definition lives next to SelectMode further down.
static AppKind ClassifyByExtension(LPCWSTR path);

// Refreshes the mode-switch button. Called after every LOAD and every
// resize. Hides the button when the current file type doesn't support
// the alternative mode (i.e. when neither Word, Excel nor PowerPoint
// recognise the extension — no quick→full switch makes sense for those).
static void UpdateModeButtonSta()
{
    if (!g_state.hwndModeButton) return;

    // Show the button only when the INI-configured mode for this file's
    // application type is one of the switchable variants.  Files whose
    // type is Other (MSG, VSDX, …) return Mode::Quick from SelectMode,
    // which is not switchable, so they naturally hide the button.
    const bool showable = IsSwitchable(SelectMode(g_state.currentFileApp));
    if (!showable || !g_state.hwndRender || !IsWindow(g_state.hwndRender))
    {
        ShowWindow(g_state.hwndModeButton, SW_HIDE);
        return;
    }

    // The label tells the user what clicking will do — i.e. the *target*
    // mode, not the current one.
    LPCWSTR label = (g_state.currentLoadedMode == Mode::Quick) ? L"→ Full"
                                                                : L"→ Quick";
    SetWindowTextW(g_state.hwndModeButton, label);

    UINT dpi = GetDpiForWindow(g_state.hwndRender);
    int w   = ScaleForDpi(64, dpi);
    int h   = ScaleForDpi(24, dpi);
    int pad = ScaleForDpi(4,  dpi);

    RECT rc; GetClientRect(g_state.hwndRender, &rc);
    // Top-right corner of the render pane. HWND_TOP keeps the button above
    // any sibling that the preview handler or Office reparented in.
    SetWindowPos(g_state.hwndModeButton, HWND_TOP,
                 (rc.right - rc.left) - w - pad, pad, w, h,
                 SWP_NOACTIVATE | SWP_SHOWWINDOW);
}

// Create our local render window and reparent it into the plugin's child
// window (which lives inside Total Commander's process). Cross-process
// SetParent is supported and is the standard pattern for hosting shell
// extensions in foreign containers.
static bool CreateRenderWindowSta(HINSTANCE hInst)
{
    EnsureRenderClass(hInst);

    RECT rc = {};
    GetClientRect(g_state.hwndPluginChild, &rc);
    const int w = rc.right - rc.left;
    const int h = rc.bottom - rc.top;
    HostLog(L"CreateRenderWindowSta: plugin child rect = %dx%d", w, h);

    // Create as message-only first so the window is born already detached
    // from a real visual tree, then reparent into the plugin's container.
    HWND hRender = CreateWindowExW(
        0, kRenderClassName, L"",
        WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        0, 0, w, h,
        HWND_MESSAGE, nullptr, hInst, nullptr);
    if (!hRender)
    {
        HostLog(L"  CreateWindowEx failed err=%lu", GetLastError());
        return false;
    }
    HostLog(L"  render hwnd = 0x%p", hRender);

    if (!SetParent(hRender, g_state.hwndPluginChild))
    {
        HostLog(L"  SetParent failed err=%lu", GetLastError());
        DestroyWindow(hRender);
        return false;
    }
    HostLog(L"  SetParent into 0x%p OK", g_state.hwndPluginChild);

    // SetParent does not modify styles automatically, but it does drop the
    // WS_POPUP bit and add WS_CHILD if needed. Make sure visibility and
    // clipping are right, then show.
    SetWindowLongPtrW(hRender, GWL_STYLE,
                      WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE);
    SetWindowPos(hRender, nullptr, 0, 0, w, h,
                 SWP_NOZORDER | SWP_FRAMECHANGED | SWP_SHOWWINDOW);

    g_state.hwndRender = hRender;

    // Mode-switch overlay button — child of the render window, stays hidden
    // until the first LOAD picks an AppKind that supports both modes.
    CreateModeButtonSta();

    // Periodic Z-order keep-alive: some preview handlers (Word in particular)
    // call SetWindowPos on their own window during scrolling without
    // SWP_NOZORDER, silently moving themselves above the mode button.
    // WM_SIZE is never fired in that case, so the one-shot resize timer
    // cannot catch it.  A low-frequency periodic timer handles it instead.
    SetTimer(g_state.hwndRender, kModeButtonKeepTopTimerId, 100, nullptr);

    return true;
}

// Forward declaration — defined later in the close-guard section.
static void DestroyCloseGuard();

static void DestroyRenderWindowSta()
{
    DestroyCloseGuard();   // must go before hwndRender is destroyed
    if (g_state.hwndRender)
    {
        // Cancel both timers before the window dies.
        KillTimer(g_state.hwndRender, kModeButtonZTimerId);
        KillTimer(g_state.hwndRender, kModeButtonKeepTopTimerId);
        // Reparent back to HWND_MESSAGE before destroying — this severs the
        // cross-process link cleanly so TC's UI thread doesn't see a stale
        // child reference while we're tearing down.
        SetParent(g_state.hwndRender, HWND_MESSAGE);
        DestroyWindow(g_state.hwndRender);
        g_state.hwndRender = nullptr;
    }
    g_state.hwndModeButton = nullptr;            // was a child of hwndRender
    if (g_state.hModeButtonFont)
    {
        DeleteObject(g_state.hModeButtonFont);
        g_state.hModeButtonFont = nullptr;
    }
}

// ---------------------------------------------------------------------------
// Fallback UI
//
// Displayed inside the render window whenever a real preview is not
// available — most commonly because MS Office is not installed (so no
// preview handler is registered for the file extension), but also when a
// registered handler fails to load. A read-only multi-line EDIT control
// shows an explanatory message plus whatever metadata we can extract from
// the file system, so the user still gets something useful in the Lister
// pane instead of an empty panel.
// ---------------------------------------------------------------------------

static int CALLBACK FontExistsCallback(const LOGFONTW*, const TEXTMETRICW*,
                                       DWORD, LPARAM lp)
{
    *reinterpret_cast<bool*>(lp) = true;
    return 0;   // stop enumerating — one hit is all we need
}

static bool FontInstalled(LPCWSTR family)
{
    HDC hdc = GetDC(nullptr);
    if (!hdc) return false;
    LOGFONTW lf = {};
    lf.lfCharSet = DEFAULT_CHARSET;
    wcscpy_s(lf.lfFaceName, ARRAYSIZE(lf.lfFaceName), family);
    bool found = false;
    EnumFontFamiliesExW(hdc, &lf, FontExistsCallback,
                        reinterpret_cast<LPARAM>(&found), 0);
    ReleaseDC(nullptr, hdc);
    return found;
}

// Pick the first monospace face that actually exists on this system.
// Order: Microsoft's most recent, modern, polished offering first; then
// progressively older / more boring options. The final entry is "Courier
// New" which has shipped with every Windows since 3.x — guaranteed.
static LPCWSTR AutoPickFontFamily()
{
    static const wchar_t* kCandidates[] = {
        L"Aptos Mono",      // Microsoft 365 / Office 2024
        L"Consolas",        // ships with Windows Vista+
        L"Cascadia Mono",   // ships with newer Windows / Windows Terminal
        L"Lucida Console",  // ships with Windows 2000+
    };
    for (LPCWSTR name : kCandidates)
        if (FontInstalled(name)) return name;
    return L"Courier New";
}

// Build a fresh HFONT sized for the actual DPI of the target window. We
// recreate it on every ShowFallbackSta call so that moving the Lister
// across monitors between previews picks up the new DPI.
static HFONT CreateFallbackFont(HWND hwnd)
{
    UINT dpi = GetDpiForWindow(hwnd);
    if (dpi == 0) dpi = USER_DEFAULT_SCREEN_DPI;   // 96
    int pixelHeight = -MulDiv(g_fontSize, dpi, USER_DEFAULT_SCREEN_DPI);

    LPCWSTR family = g_fontFamily.empty()
        ? AutoPickFontFamily()
        : g_fontFamily.c_str();

    HostLog(L"  CreateFallbackFont: family='%s' size=%dpt dpi=%u",
            family, g_fontSize, dpi);

    return CreateFontW(pixelHeight, 0, 0, 0,
                       FW_NORMAL, FALSE, FALSE, FALSE,
                       DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
                       CLEARTYPE_QUALITY, FIXED_PITCH | FF_MODERN, family);
}

static std::wstring FormatFileSize(ULONGLONG bytes)
{
    wchar_t scaled[32] = {};
    if (bytes < 1024ULL)
        _snwprintf_s(scaled, _TRUNCATE, L"%llu B", bytes);
    else if (bytes < 1024ULL * 1024)
        _snwprintf_s(scaled, _TRUNCATE, L"%.1f KB", bytes / 1024.0);
    else if (bytes < 1024ULL * 1024 * 1024)
        _snwprintf_s(scaled, _TRUNCATE, L"%.1f MB", bytes / (1024.0 * 1024));
    else
        _snwprintf_s(scaled, _TRUNCATE, L"%.2f GB", bytes / (1024.0 * 1024 * 1024));

    wchar_t raw[48] = {};
    _snwprintf_s(raw, _TRUNCATE, L"  (%llu bytes)", bytes);
    return std::wstring(scaled) + raw;
}

static std::wstring FormatFileTime(const FILETIME& ft)
{
    if (ft.dwLowDateTime == 0 && ft.dwHighDateTime == 0) return L"(not set)";
    FILETIME local = {};
    if (!FileTimeToLocalFileTime(&ft, &local)) return L"(?)";
    SYSTEMTIME st = {};
    if (!FileTimeToSystemTime(&local, &st)) return L"(?)";
    wchar_t buf[64] = {};
    _snwprintf_s(buf, _TRUNCATE, L"%04u-%02u-%02u %02u:%02u:%02u",
                 st.wYear, st.wMonth, st.wDay,
                 st.wHour, st.wMinute, st.wSecond);
    return buf;
}

// ---------------------------------------------------------------------------
// Mark of the Web (MOTW) detection and removal.
//
// Files downloaded from the internet carry an alternate data stream named
// "Zone.Identifier" with an INI-shaped body identifying the security zone:
//
//   [ZoneTransfer]
//   ZoneId=3
//   ReferrerUrl=https://…
//   HostUrl=https://…
//
// ZoneId values per Microsoft Internet zones:
//   0 = Local Computer
//   1 = Local Intranet
//   2 = Trusted Sites
//   3 = Internet
//   4 = Restricted Sites
//
// Microsoft Office's Protected View blocks rendering of files in zone 3 or
// higher until the user explicitly approves them.  When that happens
// IPreviewHandler::DoPreview returns E_FAIL and our fallback panel takes
// over — we then read the Zone.Identifier ADS to tell the user precisely
// why the preview failed.
//
// Deleting the ADS removes the mark (no admin rights required) and is the
// programmatic equivalent of clicking "Enable Editing" in Office's yellow
// security banner.  Irreversible from our side — once unblocked, Windows
// treats the file as locally created and Office trusts it from then on.
// ---------------------------------------------------------------------------

struct MotwInfo
{
    int          zoneId = -1;     // -1 = no MOTW present
    std::wstring referrerUrl;
    std::wstring hostUrl;
};

// Convert a narrow string slice to wide assuming UTF-8.  Used for the
// URLs read out of the Zone.Identifier ADS.
static std::wstring NarrowToWide(const char* s, size_t n)
{
    if (n == 0) return L"";
    int needed = MultiByteToWideChar(CP_UTF8, 0, s, static_cast<int>(n),
                                     nullptr, 0);
    if (needed <= 0) return L"";
    std::wstring w(static_cast<size_t>(needed), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s, static_cast<int>(n),
                        w.data(), needed);
    return w;
}

static MotwInfo ReadFileZoneInfo(LPCWSTR path)
{
    MotwInfo info;

    const std::wstring adsPath = std::wstring(path) + L":Zone.Identifier";
    HANDLE h = CreateFileW(adsPath.c_str(), GENERIC_READ,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           nullptr, OPEN_EXISTING,
                           FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return info;       // no MOTW

    char buf[4096] = {};
    DWORD bytesRead = 0;
    BOOL ok = ReadFile(h, buf, sizeof(buf) - 1, &bytesRead, nullptr);
    CloseHandle(h);
    if (!ok || bytesRead == 0) return info;
    buf[bytesRead] = '\0';

    // Pulls "<key>=<value>" out of the ADS body, value bounded by EOL.
    auto extract = [&](const char* key) -> std::wstring {
        const char* p = strstr(buf, key);
        if (!p) return L"";
        p += strlen(key);
        const char* end = p;
        while (*end && *end != '\r' && *end != '\n') ++end;
        // Trim ASCII whitespace.
        while (p < end && (*p == ' ' || *p == '\t')) ++p;
        while (end > p && (end[-1] == ' ' || end[-1] == '\t')) --end;
        return NarrowToWide(p, static_cast<size_t>(end - p));
    };

    std::wstring zoneStr = extract("ZoneId=");
    if (!zoneStr.empty()) info.zoneId = _wtoi(zoneStr.c_str());
    info.referrerUrl = extract("ReferrerUrl=");
    info.hostUrl     = extract("HostUrl=");
    return info;
}

// Human-readable name for a Windows Internet Security zone.
static LPCWSTR ZoneIdName(int zoneId)
{
    switch (zoneId)
    {
        case 0: return L"Local Computer";
        case 1: return L"Local Intranet";
        case 2: return L"Trusted Sites";
        case 3: return L"Internet";
        case 4: return L"Restricted / Untrusted Sites";
        default: return L"(unknown zone)";
    }
}

// Strip the Zone.Identifier ADS from a file.  Equivalent to PowerShell
// `Unblock-File` or right-click → Properties → Unblock.  Returns true
// on success.  No-op (but returns true) if the ADS is absent.
static bool UnblockFileSta(LPCWSTR path)
{
    const std::wstring adsPath = std::wstring(path) + L":Zone.Identifier";
    if (DeleteFileW(adsPath.c_str())) return true;
    DWORD err = GetLastError();
    // ERROR_FILE_NOT_FOUND means there was nothing to remove — that's
    // still "success" from the caller's point of view.
    return err == ERROR_FILE_NOT_FOUND;
}

static std::wstring BuildFallbackText(LPCWSTR path, HRESULT hr)
{
    std::wstring text;

    // Mark of the Web is by far the most common reason an otherwise-
    // healthy Office handler refuses to render — check first so the
    // user gets a precise diagnosis instead of a generic "handler
    // failed" message.
    const MotwInfo motw = ReadFileZoneInfo(path);
    const bool isMotwBlocked = (motw.zoneId >= 3);   // Internet / Untrusted

    if (isMotwBlocked)
    {
        text += L"This file is blocked because Windows marked it as coming\r\n"
                L"from outside your computer — downloaded from the internet,\r\n"
                L"copied from a network share, or saved from an email\r\n"
                L"attachment.\r\n"
                L"\r\n"
                L"Microsoft Office refuses to render such files until you\r\n"
                L"explicitly trust them.  This protects you from malicious\r\n"
                L"macros, exploits and weaponised documents.  Unblock only\r\n"
                L"files whose source you trust.\r\n"
                L"\r\n"
                L"Source of the file\r\n";

        wchar_t zoneLine[128] = {};
        _snwprintf_s(zoneLine, _TRUNCATE,
                     L"  Zone:        %s (ZoneId=%d)\r\n",
                     ZoneIdName(motw.zoneId), motw.zoneId);
        text += zoneLine;

        if (!motw.hostUrl.empty())
        {
            text += L"  Came from:   ";
            text += motw.hostUrl;
            text += L"\r\n";
        }
        if (!motw.referrerUrl.empty() && motw.referrerUrl != motw.hostUrl)
        {
            text += L"  Referrer:    ";
            text += motw.referrerUrl;
            text += L"\r\n";
        }

        text += L"\r\n"
                L"Clicking \"Unblock\" below will:\r\n"
                L"  - Remove the download mark from this file (this file\r\n"
                L"    only).\r\n"
                L"  - Retry the preview.\r\n"
                L"  - Permanently mark the file as trusted on this\r\n"
                L"    computer; Office will open it without warning from\r\n"
                L"    now on.\r\n"
                L"\r\n"
                L"This is equivalent to clicking \"Enable Editing\" in\r\n"
                L"Office's yellow security bar.  If you are unsure, close\r\n"
                L"this preview and inspect the file with another tool\r\n"
                L"first.\r\n";
    }
    else if (hr == REGDB_E_CLASSNOTREG)
    {
        text += L"No MS Office preview handler is registered for this file.\r\n";
        text += L"\r\n";
        text += L"TCOfficeView renders previews by hosting the Windows Preview\r\n";
        text += L"Handlers that Microsoft Office registers when it is installed.\r\n";
        text += L"No such handler was found on this computer for this file type.\r\n";
        text += L"\r\n";
        text += L"To enable full document previews, install Microsoft Office\r\n";
        text += L"(any edition that includes the relevant application — Word\r\n";
        text += L"for DOC/DOCX, Excel for XLS/XLSX, PowerPoint for PPT/PPTX,\r\n";
        text += L"Outlook for MSG, Visio for VSD/VSDX, ...).\r\n";
    }
    else
    {
        wchar_t lead[256] = {};
        _snwprintf_s(lead, _TRUNCATE,
                     L"This file could not be previewed.\r\n"
                     L"\r\n"
                     L"The registered preview handler failed to load it\r\n"
                     L"(HRESULT 0x%08lX).\r\n",
                     static_cast<long>(hr));
        text += lead;

        const AppKind app = ClassifyByExtension(path);
        const bool isOfficeApp = (app == AppKind::Word ||
                                  app == AppKind::Excel ||
                                  app == AppKind::PowerPoint);

        if (isOfficeApp)
        {
            text += L"\r\n"
                    L"Common causes for Office files:\r\n"
                    L"  - The document is on a SharePoint site synced from a\r\n"
                    L"    Microsoft 365 tenant other than your primary one —\r\n"
                    L"    the sandboxed preview-handler process cannot\r\n"
                    L"    authenticate cross-tenant.  Files affected by this\r\n"
                    L"    also refuse to preview in Windows Explorer's Alt+P\r\n"
                    L"    pane; it is an Office limitation, not a\r\n"
                    L"    TCOfficeView one.\r\n"
                    L"  - The document is corrupted or uses a format the\r\n"
                    L"    handler does not recognise.\r\n"
                    L"  - The Office installation is broken.\r\n";

            // The "→ Full" hint only makes sense when the mode-switch
            // button is actually on screen.  IsSwitchable() returns false
            // for the explicit `quick` (non-switchable) configuration —
            // the user has opted out of mode switching and the button
            // is hidden, so pointing at it would just confuse.
            if (IsSwitchable(SelectMode(app)))
            {
                text += L"\r\n"
                        L"Tip: click the \"→ Full\" button in the top-right\r\n"
                        L"corner of this pane to open the document in the real\r\n"
                        L"Office application instead.  It runs as your user\r\n"
                        L"with full credentials and can usually open files the\r\n"
                        L"preview handler refuses.\r\n";
            }
        }
        else
        {
            text += L"\r\n"
                    L"The application that registered the handler may be\r\n"
                    L"installed but broken, or the file itself may be\r\n"
                    L"corrupted.\r\n";
        }
    }

    text += L"\r\n";
    text += L"--------------------------------------------------------------\r\n";
    text += L"File information\r\n";
    text += L"\r\n";

    LPCWSTR fname = wcsrchr(path, L'\\');
    fname = fname ? fname + 1 : path;
    text += L"  Name:        ";
    text += fname;
    text += L"\r\n";
    text += L"  Path:        ";
    text += path;
    text += L"\r\n";

    LPCWSTR ext = wcsrchr(path, L'.');
    if (ext && ext[1])
    {
        text += L"  Extension:   ";
        text += ext;
        text += L"\r\n";
    }

    // Use the long-path-safe form when querying file attributes; the
    // panel still displays the user's original path below.
    const std::wstring longPath = EnsureLongPathPrefix(path);
    WIN32_FILE_ATTRIBUTE_DATA fad = {};
    if (GetFileAttributesExW(longPath.c_str(), GetFileExInfoStandard, &fad))
    {
        ULARGE_INTEGER size;
        size.HighPart = fad.nFileSizeHigh;
        size.LowPart  = fad.nFileSizeLow;
        text += L"  Size:        ";
        text += FormatFileSize(size.QuadPart);
        text += L"\r\n";
        text += L"  Modified:    ";
        text += FormatFileTime(fad.ftLastWriteTime);
        text += L"\r\n";
        text += L"  Created:     ";
        text += FormatFileTime(fad.ftCreationTime);
        text += L"\r\n";
        text += L"  Accessed:    ";
        text += FormatFileTime(fad.ftLastAccessTime);
        text += L"\r\n";
    }
    else
    {
        wchar_t err[64] = {};
        _snwprintf_s(err, _TRUNCATE,
                     L"  (file attributes could not be read, err=%lu)\r\n",
                     GetLastError());
        text += err;
    }

    return text;
}

static void HideFallbackSta()
{
    if (g_state.hwndFallback)
    {
        DestroyWindow(g_state.hwndFallback);
        g_state.hwndFallback = nullptr;
    }
    if (g_state.hFallbackFont)
    {
        DeleteObject(g_state.hFallbackFont);
        g_state.hFallbackFont = nullptr;
    }
    if (g_state.hwndUnblockButton)
    {
        DestroyWindow(g_state.hwndUnblockButton);
        g_state.hwndUnblockButton = nullptr;
    }
    if (g_state.hUnblockButtonFont)
    {
        DeleteObject(g_state.hUnblockButtonFont);
        g_state.hUnblockButtonFont = nullptr;
    }
}

// Position the MOTW unblock button at the bottom-centre of the render pane.
// No-op if the button does not currently exist.  Centralised because both
// ShowFallbackSta (initial placement) and ResizeHandlerSta (on every WM_SIZE)
// need the exact same geometry — without the resize path the button stays at
// its original coordinates and can end up off-screen when the user shrinks
// the Lister window.
static void LayoutUnblockButtonSta(int paneW, int paneH)
{
    if (!g_state.hwndUnblockButton) return;
    UINT dpi = GetDpiForWindow(g_state.hwndRender);
    if (dpi == 0) dpi = USER_DEFAULT_SCREEN_DPI;
    const int btnHeight  = ScaleForDpi(36, dpi);
    const int btnPadding = ScaleForDpi(8, dpi);
    const int btnWidth   = ScaleForDpi(320, dpi);
    const int btnX = (paneW - btnWidth) / 2;
    const int btnY = paneH - btnHeight - btnPadding;
    // HWND_TOP keeps it above any sibling the close guard or a leftover
    // Office frame may have added to hwndRender's child list.
    SetWindowPos(g_state.hwndUnblockButton, HWND_TOP,
                 btnX, btnY, btnWidth, btnHeight,
                 SWP_NOACTIVATE);
}

static bool ShowFallbackSta(LPCWSTR path, HRESULT hr)
{
    HideFallbackSta();
    if (!g_state.hwndRender) return false;

    RECT rc = {};
    GetClientRect(g_state.hwndRender, &rc);

    const int paneW = rc.right - rc.left;
    const int paneH = rc.bottom - rc.top;

    // Decide whether to reserve space at the bottom for an "Unblock"
    // button.  Only when the file is actually MOTW-blocked do we want
    // to offer the action; for plain "no handler" / "handler crashed"
    // fallbacks the button would have nothing useful to do.
    const MotwInfo motw = ReadFileZoneInfo(path);
    const bool showUnblock = (motw.zoneId >= 3);

    UINT dpi = GetDpiForWindow(g_state.hwndRender);
    if (dpi == 0) dpi = USER_DEFAULT_SCREEN_DPI;

    // The EDIT covers the FULL pane regardless of whether we add a
    // button.  This is intentional: when navigating between files the
    // previous preview handler often leaves its rendering window
    // parented under hwndRender (the surrogate process keeps it alive
    // even after our COM Unload + Release).  A smaller EDIT would let
    // those leftover windows show through in the side / bottom gaps
    // around an undersized button.  Making the EDIT full-pane means
    // its opaque background occludes every leftover sibling cleanly;
    // the Unblock button is then placed ON TOP, overlaying the EDIT
    // at the bottom — also opaque, so its area is clean too.  The
    // last few lines of the fallback text may be hidden behind the
    // button; the EDIT's vertical scroll bar lets the user reach them
    // and the important content is in the upper paragraphs anyway.
    HWND hEdit = CreateWindowExW(
        0, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_CLIPSIBLINGS |
        ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL | ES_LEFT,
        0, 0, paneW, paneH,
        g_state.hwndRender, nullptr,
        GetModuleHandleW(nullptr), nullptr);
    if (!hEdit)
    {
        HostLog(L"  ShowFallbackSta: CreateWindowEx(EDIT) failed err=%lu",
                GetLastError());
        return false;
    }

    g_state.hFallbackFont = CreateFallbackFont(g_state.hwndRender);
    SendMessageW(hEdit, WM_SETFONT,
                 reinterpret_cast<WPARAM>(g_state.hFallbackFont), TRUE);

    std::wstring body = BuildFallbackText(path, hr);
    SetWindowTextW(hEdit, body.c_str());

    g_state.hwndFallback = hEdit;

    // Unblock button (MOTW only).  Posts WM_HOST_UNBLOCK_AND_RELOAD on
    // click; the STA handler strips the file's Zone.Identifier ADS and
    // re-runs the LOAD.
    //
    // BS_OWNERDRAW so RenderWndProc's WM_DRAWITEM handler can paint it
    // in an attention-grabbing dark red — the default themed grey blends
    // into the fallback panel and users miss it.  On owner-draw push
    // buttons WM_CTLCOLORBTN is NOT sent on themed Windows, so this is
    // the only reliable way to colorise.
    if (showUnblock)
    {
        wchar_t btnLabel[128] = {};
        _snwprintf_s(btnLabel, _TRUNCATE,
                     L"Unblock this file and retry the preview  (%s)",
                     ZoneIdName(motw.zoneId));

        HWND hBtn = CreateWindowExW(
            0, L"BUTTON", btnLabel,
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | BS_OWNERDRAW,
            0, 0, 0, 0,                              // sized by LayoutUnblockButtonSta
            g_state.hwndRender,
            reinterpret_cast<HMENU>(kUnblockButtonId),
            GetModuleHandleW(nullptr), nullptr);
        if (hBtn)
        {
            int fontHeight = -ScaleForDpi(11, dpi);   // ~11pt at the window's DPI
            HFONT hFont = CreateFontW(
                fontHeight, 0, 0, 0,
                FW_BOLD, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
            if (hFont)
                SendMessageW(hBtn, WM_SETFONT,
                             reinterpret_cast<WPARAM>(hFont), TRUE);
            g_state.hwndUnblockButton  = hBtn;
            g_state.hUnblockButtonFont = hFont;
            LayoutUnblockButtonSta(paneW, paneH);
            HostLog(L"  ShowFallbackSta: MOTW detected (ZoneId=%d), "
                    L"Unblock button shown", motw.zoneId);
        }
        else
        {
            HostLog(L"  ShowFallbackSta: MOTW detected but Button "
                    L"CreateWindowEx failed err=%lu", GetLastError());
        }
    }
    // Refresh the mode-switch button so it is re-positioned above the
    // fallback controls (the EDIT was just created and sits at the top
    // of the Z-order by default).
    UpdateModeButtonSta();

    // Force a full repaint of the render pane and every child.  Without
    // this, a previous handler's leftover bitmap can briefly bleed through
    // before the OS gets around to invalidating the area newly covered by
    // our EDIT / button.
    RedrawWindow(g_state.hwndRender, nullptr, nullptr,
                 RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);

    HostLog(L"  ShowFallbackSta: displayed fallback UI (hr=0x%08lX)",
            static_cast<long>(hr));
    return true;
}

// ---------------------------------------------------------------------------
// Application dispatch (extension → app → mode)
// ---------------------------------------------------------------------------

static AppKind ClassifyByExtension(LPCWSTR path)
{
    LPCWSTR dot = wcsrchr(path, L'.');
    if (!dot || !dot[1]) return AppKind::Other;
    // Compare case-insensitively against the known extensions.
    struct Entry { const wchar_t* ext; AppKind app; };
    static const Entry kTable[] = {
        { L".doc",  AppKind::Word },       { L".docx", AppKind::Word },
        { L".docm", AppKind::Word },       { L".rtf",  AppKind::Word },
        { L".xls",  AppKind::Excel },      { L".xlsx", AppKind::Excel },
        { L".xlsm", AppKind::Excel },      { L".xlsb", AppKind::Excel },
        { L".ppt",  AppKind::PowerPoint }, { L".pptx", AppKind::PowerPoint },
        { L".pptm", AppKind::PowerPoint },
    };
    for (const auto& e : kTable)
        if (_wcsicmp(dot, e.ext) == 0) return e.app;
    return AppKind::Other;
}

static Mode SelectMode(AppKind app)
{
    switch (app)
    {
        case AppKind::Word:       return g_modeWord;
        case AppKind::Excel:      return g_modeExcel;
        case AppKind::PowerPoint: return g_modePowerPoint;
        default:                  return Mode::Quick;
    }
}

static bool IsAutoFallbackEnabled(AppKind app)
{
    switch (app)
    {
        case AppKind::Word:       return g_autoFallbackWord;
        case AppKind::Excel:      return g_autoFallbackExcel;
        case AppKind::PowerPoint: return g_autoFallbackPowerPoint;
        default:                  return false;
    }
}

// ---------------------------------------------------------------------------
// IDispatch helpers — thin wrappers around the standard OLE Automation
// dance (GetIDsOfNames → Invoke). All take/return owned VARIANTs; callers
// must VariantClear what they receive.
// ---------------------------------------------------------------------------

static HRESULT DispGetId(IDispatch* p, LPCOLESTR name, DISPID* outId)
{
    LPOLESTR n = const_cast<LPOLESTR>(name);
    return p->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, outId);
}

// Invoke a method or property-get on an IDispatch with `cArgs` positional
// arguments. Args are passed in REVERSE order, as required by IDispatch::Invoke.
static HRESULT DispCall(IDispatch* p, LPCOLESTR name, WORD flags,
                        VARIANT* args, UINT cArgs, VARIANT* outResult)
{
    DISPID id = 0;
    HRESULT hr = DispGetId(p, name, &id);
    if (FAILED(hr)) return hr;

    DISPPARAMS dp = {};
    dp.cArgs = cArgs;
    dp.rgvarg = args;

    // Property-put requires DISPID_PROPERTYPUT named arg.
    DISPID putDispId = DISPID_PROPERTYPUT;
    if (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF))
    {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &putDispId;
    }

    EXCEPINFO ex = {};
    UINT argErr = 0;
    return p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                     &dp, outResult, &ex, &argErr);
}

static HRESULT DispGetProperty(IDispatch* p, LPCOLESTR name, VARIANT* out)
{
    VariantInit(out);
    return DispCall(p, name, DISPATCH_PROPERTYGET, nullptr, 0, out);
}

static HRESULT DispPutBool(IDispatch* p, LPCOLESTR name, bool value)
{
    VARIANT v;  VariantInit(&v);
    v.vt = VT_BOOL;
    v.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    return DispCall(p, name, DISPATCH_PROPERTYPUT, &v, 1, nullptr);
}

static HRESULT DispPutI4(IDispatch* p, LPCOLESTR name, LONG value)
{
    VARIANT v;  VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = value;
    return DispCall(p, name, DISPATCH_PROPERTYPUT, &v, 1, nullptr);
}

// Read an IDispatch property whose return type is another IDispatch (used
// to walk Office object chains, e.g. Application.ActiveWindow.View).
// Returns nullptr on failure or wrong VT; caller owns and must Release().
static IDispatch* DispGetDispatchProperty(IDispatch* p, LPCOLESTR name)
{
    VARIANT v;
    if (FAILED(DispGetProperty(p, name, &v))) return nullptr;
    if (v.vt != VT_DISPATCH || !v.pdispVal)
    {
        VariantClear(&v);
        return nullptr;
    }
    return v.pdispVal;       // ownership transferred to caller
}

// Read an HWND from an IDispatch property whose return type is some flavour
// of integer (Word/Excel/PowerPoint all expose Application.Hwnd this way,
// though the exact VT varies between Office versions).
static HWND DispGetHwndProperty(IDispatch* p, LPCOLESTR name)
{
    VARIANT v;
    HRESULT hr = DispGetProperty(p, name, &v);
    if (FAILED(hr))
    {
        HostLog(L"  DispGetHwndProperty(%s): GetProperty -> 0x%08lX",
                name, static_cast<long>(hr));
        return nullptr;
    }
    HWND h = nullptr;
    switch (v.vt)
    {
        case VT_I4:   h = reinterpret_cast<HWND>(static_cast<INT_PTR>(v.lVal));    break;
        case VT_I8:   h = reinterpret_cast<HWND>(static_cast<INT_PTR>(v.llVal));   break;
        case VT_INT:  h = reinterpret_cast<HWND>(static_cast<INT_PTR>(v.intVal));  break;
        case VT_UI4:  h = reinterpret_cast<HWND>(static_cast<UINT_PTR>(v.ulVal));  break;
        case VT_UI8:  h = reinterpret_cast<HWND>(static_cast<UINT_PTR>(v.ullVal)); break;
        case VT_UINT: h = reinterpret_cast<HWND>(static_cast<UINT_PTR>(v.uintVal));break;
        default:
            HostLog(L"  DispGetHwndProperty(%s): unexpected VT=%d",
                    name, static_cast<int>(v.vt));
            break;
    }
    VariantClear(&v);
    return h;
}

// Fallback HWND finder used when Application.Hwnd refuses to give us one.
// Walks all top-level windows looking for the requested Office frame class
// ("OpusApp" for Word, "XLMAIN" for Excel, "PPTFrameClass" for PowerPoint),
// optionally filtered by a substring of the window title (typically the
// filename of the document we just opened, so we don't pick up an instance
// the user already had running for a different document).
struct FindOfficeWndCtx
{
    LPCWSTR  className;
    LPCWSTR  titleSubstr;       // may be null
    HWND     found;
};

static BOOL CALLBACK FindOfficeWndEnumProc(HWND hwnd, LPARAM lp)
{
    auto* ctx = reinterpret_cast<FindOfficeWndCtx*>(lp);
    wchar_t cls[64] = {};
    if (GetClassNameW(hwnd, cls, ARRAYSIZE(cls)) <= 0) return TRUE;
    if (wcscmp(cls, ctx->className) != 0) return TRUE;
    if (ctx->titleSubstr)
    {
        wchar_t title[512] = {};
        GetWindowTextW(hwnd, title, ARRAYSIZE(title));
        if (!StrStrIW(title, ctx->titleSubstr)) return TRUE;
    }
    ctx->found = hwnd;
    return FALSE;                          // stop enumerating
}

static HWND FindOfficeTopLevelWindow(LPCWSTR className, LPCWSTR titleSubstr)
{
    FindOfficeWndCtx ctx = { className, titleSubstr, nullptr };
    EnumWindows(FindOfficeWndEnumProc, reinterpret_cast<LPARAM>(&ctx));
    return ctx.found;
}

// ---------------------------------------------------------------------------
// Full-mode Word lifecycle. STA thread only.
//
// We drive a Word.Application instance via OLE Automation and reparent its
// top-level main window into our render pane. The instance is created lazily
// on the first full-mode LOAD and kept alive across subsequent LOADs within
// the same Lister session (matching the persistence behaviour of the quick-
// mode handler). On CLOSE we Application.Quit and release.
//
// Word notoriously dislikes being driven for "preview" use cases — modal
// dialogs (AutoRecovery, Activation, Trust Center) can pop up if Word's
// global state is broken. We set DisplayAlerts=wdAlertsNone to suppress
// the common ones, but a deeply broken install will still cause trouble.
// ---------------------------------------------------------------------------

// Forward declarations — each LoadXxxFullSta tears down the quick-mode
// handler and the other two Office apps before installing its own window
// into hwndRender. Definitions live in the sections below.
static void UnloadHandlerSta();
static void UnloadWordFullSta(bool quitApp);
static void UnloadExcelFullSta(bool quitApp);
static void UnloadPptFullSta(bool quitApp);

static void CloseWordDocumentSta()
{
    if (!g_state.pWordDoc) return;
    VARIANT vSaveChanges; VariantInit(&vSaveChanges);
    vSaveChanges.vt = VT_I4;
    vSaveChanges.lVal = 0;                  // wdDoNotSaveChanges
    DispCall(g_state.pWordDoc, L"Close", DISPATCH_METHOD, &vSaveChanges, 1, nullptr);
    g_state.pWordDoc->Release();
    g_state.pWordDoc = nullptr;
}

// Detach Word's main window from our render pane back to top-level and hide
// it. Called when switching out of full-Word mode (different file type) or
// before Application.Quit. We don't destroy the HWND — Word owns it; we just
// undo the SetParent/style edits we did at load time.
static void DetachWordWindowSta()
{
    if (!g_state.hwndWordApp) return;
    ShowWindow(g_state.hwndWordApp, SW_HIDE);
    DestroyCloseGuard();
    SetParent(g_state.hwndWordApp, nullptr);
    LONG_PTR style = GetWindowLongPtrW(g_state.hwndWordApp, GWL_STYLE);
    style &= ~WS_CHILD;
    style |= WS_OVERLAPPEDWINDOW;
    SetWindowLongPtrW(g_state.hwndWordApp, GWL_STYLE, style);
    // Force the window manager to re-evaluate the non-client area
    // (caption, borders) now that the style changed back to overlapped.
    // Without this the frame can stay in an inconsistent state until
    // the next paint, causing ghost windows or focus glitches.
    SetWindowPos(g_state.hwndWordApp, nullptr, 0, 0, 0, 0,
                 SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE |
                 SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);
    g_state.hwndWordApp = nullptr;
}

static void UnloadWordFullSta(bool quitApp)
{
    CloseWordDocumentSta();
    DetachWordWindowSta();
    if (quitApp && g_state.pWordApp)
    {
        DispCall(g_state.pWordApp, L"Quit", DISPATCH_METHOD, nullptr, 0, nullptr);
        g_state.pWordApp->Release();
        g_state.pWordApp = nullptr;
    }
}

// Tweak Word's UI for preview use. Strict rule: we only touch state that is
// scoped to the active window / view / document. We never modify any
// Application-level property (DisplayStatusBar, DisplayRibbon, the
// CommandBars.ExecuteMso "MinimizeRibbon" toggle, …) because Office
// persists those into the user's profile on Quit, and the user's regular
// standalone Word would then start with our preview-friendly settings
// instead of theirs.
//
// All steps are best-effort — any failure is logged but does not abort the
// load, because the embedded view still works without these tweaks.
static void ConfigureWordForPreviewSta()
{
    // ActiveWindow.View.Type = wdPrintView (3) — gives us the paged layout
    // with headers, footers and page breaks that quick-mode can't render.
    // Per-window setting; does not persist after the window is closed.
    IDispatch* pWin = DispGetDispatchProperty(g_state.pWordApp, L"ActiveWindow");
    if (pWin)
    {
        IDispatch* pView = DispGetDispatchProperty(pWin, L"View");
        if (pView)
        {
            HRESULT hr = DispPutI4(pView, L"Type", 3);          // wdPrintView
            HostLog(L"  View.Type = wdPrintView -> 0x%08lX", static_cast<long>(hr));

            // Fit-to-page-width. wdPageFitBestFit scales the page to match
            // the window width and re-fits automatically when the window is
            // resized — important for the narrow Quick View (Ctrl+Q) pane.
            IDispatch* pZoom = DispGetDispatchProperty(pView, L"Zoom");
            if (pZoom)
            {
                HRESULT zhr = DispPutI4(pZoom, L"PageFit", 2);  // wdPageFitBestFit
                HostLog(L"  Zoom.PageFit = wdPageFitBestFit -> 0x%08lX",
                        static_cast<long>(zhr));
                pZoom->Release();
            }
            pView->Release();
        }
        // Per-window rulers off.
        HRESULT hr = DispPutBool(pWin, L"DisplayRulers", false);
        HostLog(L"  DisplayRulers = false -> 0x%08lX", static_cast<long>(hr));
        pWin->Release();
    }

    // Document.Protect(Type:=wdAllowOnlyReading, NoReset:=False, Password:="")
    // — true read-only enforcement. Without this, opening with ReadOnly=True
    // still lets the user "Edit Anyway" and trigger a Save As dialog on Ctrl+S.
    // Per-document setting; cannot leak because we opened the document
    // ReadOnly=True so Word cannot save the protection state back to disk.
    if (g_state.pWordDoc)
    {
        VARIANT args[3];
        for (int i = 0; i < 3; ++i) VariantInit(&args[i]);
        args[0].vt = VT_BSTR; args[0].bstrVal = SysAllocString(L"");   // Password
        args[1].vt = VT_BOOL; args[1].boolVal = VARIANT_FALSE;          // NoReset
        args[2].vt = VT_I4;   args[2].lVal    = 3;                      // Type = wdAllowOnlyReading

        HRESULT hr = DispCall(g_state.pWordDoc, L"Protect",
                              DISPATCH_METHOD, args, 3, nullptr);
        HostLog(L"  Document.Protect(wdAllowOnlyReading) -> 0x%08lX",
                static_cast<long>(hr));
        SysFreeString(args[0].bstrVal);
    }
}

// Excel preview tweak: set a fixed 100% zoom (per-window, does not modify
// the workbook on disk since we opened ReadOnly). We deliberately don't
// touch Application.WindowState — for a WS_CHILD Excel frame, both
// xlNormal and xlMaximized produce visual glitches (the latter blanks
// the ribbon entirely). Layout for the initial size is handled in
// LoadExcelFullSta by setting Application.Width/Height *before*
// Workbooks.Open; ResizeOfficeFullSta keeps them in sync on resize.
//
// Two intrinsic limits we accepted as not solvable from outside Excel:
//
//   1. Excel does not relayout its child widgets (ribbon, XLDESK, sheet
//      tabs, status bar) when the embedded Win32 frame is resized
//      programmatically — neither SetWindowPos, Application.Width/Height,
//      Application.WindowState nor ActiveWindow.WindowState reliably
//      triggers it. The initial layout is correct thanks to pre-Open
//      Width/Height in LoadExcelFullSta, but subsequent Lister resizes
//      leave Excel's content anchored to the original area. Reopening
//      the Lister forces a fresh initial layout at the new size.
//
//   2. Interactive mouse input (cell selection, dragging, sheet-tab
//      clicks, most ribbon buttons) is unreliable. Excel gates much of
//      that processing on being the foreground top-level window via
//      internal GetForegroundWindow checks, and a reparented child of
//      a foreign process never is. Faking foreground (SetForegroundWindow,
//      activation hooks, synthetic WM_ACTIVATE/WM_NCACTIVATE) either
//      stole focus from Total Commander or had no effect on Excel's
//      internal checks. Word and PowerPoint do far less of this gating,
//      which is why their reparented embeds feel interactive. Users
//      who need interaction should stay in quick mode; full mode is
//      best treated as a read-only visual preview of Excel.
static void ConfigureExcelForPreviewSta()
{
    IDispatch* pWin = DispGetDispatchProperty(g_state.pExcelApp, L"ActiveWindow");
    if (pWin)
    {
        VARIANT v; VariantInit(&v);
        v.vt = VT_I4; v.lVal = 100;                 // percentage, clamped 10..400
        HRESULT hr = DispCall(pWin, L"Zoom",
                              DISPATCH_PROPERTYPUT, &v, 1, nullptr);
        HostLog(L"  Excel ActiveWindow.Zoom = 100 -> 0x%08lX",
                static_cast<long>(hr));
        pWin->Release();
    }
}

// Tell Excel its application window is W × H pixels (converted to
// points, which is the unit Application.Width/Height use). Idempotent;
// safe to call before Workbooks.Open and again on each resize.
static void ApplyExcelAppSizeSta(int wPx, int hPx)
{
    if (!g_state.pExcelApp || !g_state.hwndRender) return;
    UINT dpi = GetDpiForWindow(g_state.hwndRender);
    if (dpi == 0) dpi = USER_DEFAULT_SCREEN_DPI;

    VARIANT vW; VariantInit(&vW);
    vW.vt = VT_R8;
    vW.dblVal = static_cast<double>(wPx) * 72.0 / dpi;
    DispCall(g_state.pExcelApp, L"Width",
             DISPATCH_PROPERTYPUT, &vW, 1, nullptr);

    VARIANT vH; VariantInit(&vH);
    vH.vt = VT_R8;
    vH.dblVal = static_cast<double>(hPx) * 72.0 / dpi;
    DispCall(g_state.pExcelApp, L"Height",
             DISPATCH_PROPERTYPUT, &vH, 1, nullptr);
}

// PowerPoint preview tweaks: ZoomToFit makes PowerPoint scale the current
// slide so it fits the View window — and unlike Excel, PowerPoint *does*
// re-fit automatically when the window is resized.
static void ConfigurePowerPointForPreviewSta()
{
    IDispatch* pWin = DispGetDispatchProperty(g_state.pPptApp, L"ActiveWindow");
    if (pWin)
    {
        IDispatch* pView = DispGetDispatchProperty(pWin, L"View");
        if (pView)
        {
            // ZoomToFit accepts an MsoTriState (Long: -1 = msoTrue).
            VARIANT v; VariantInit(&v);
            v.vt = VT_I4; v.lVal = -1;
            HRESULT hr = DispCall(pView, L"ZoomToFit",
                                  DISPATCH_PROPERTYPUT, &v, 1, nullptr);
            HostLog(L"  PowerPoint View.ZoomToFit = msoTrue -> 0x%08lX",
                    static_cast<long>(hr));
            pView->Release();
        }
        pWin->Release();
    }
}

// ---------------------------------------------------------------------------
// Close-guard overlay window.
//
// Office applications run out-of-process, so SetWindowSubclass (which only
// works on HWNDs in the same process) cannot intercept messages on the
// embedded Word / Excel / PowerPoint window. Instead we create a small,
// invisible, transparent, hit-testable child window in our own process that
// sits in the top-right corner of the render pane, directly over the area
// where Office draws its own close button. Any click that lands on the guard
// is swallowed; the Office window underneath never sees it.
//
// The guard is created once per full-mode embed and destroyed when the
// Office window is detached or the render pane is torn down.
// ---------------------------------------------------------------------------

static const wchar_t* kGuardClassName = L"TCOfficeViewCloseGuard";

static LRESULT CALLBACK GuardWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_NCHITTEST:
            // Return HTCLIENT so the guard window captures the mouse click.
            // The window paints a light gray background, so the user sees
            // the guard area but the click is swallowed.
            return HTCLIENT;

        case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // Fill with the same light gray as the class background brush
            // so the guard area is clearly visible.
            RECT rc;
            GetClientRect(hWnd, &rc);
            FillRect(hdc, &rc, reinterpret_cast<HBRUSH>(GetStockObject(LTGRAY_BRUSH)));
            EndPaint(hWnd, &ps);
            return 0;
        }
    }
    return DefWindowProcW(hWnd, msg, wp, lp);
}

static void EnsureGuardClass(HINSTANCE hInst)
{
    static bool registered = false;
    if (registered) return;
    WNDCLASSW wc = {};
    wc.lpfnWndProc   = GuardWndProc;
    wc.hInstance     = hInst;
    wc.lpszClassName = kGuardClassName;
    // Use a light gray brush so the guard is visible but not intrusive.
    // GetStockObject needs no cleanup.
    wc.hbrBackground = reinterpret_cast<HBRUSH>(GetStockObject(LTGRAY_BRUSH));
    RegisterClassW(&wc);
    registered = true;
}

static void DestroyCloseGuard();   // forward declaration

static void CreateCloseGuard()
{
    DestroyCloseGuard();   // idempotent
    if (!g_state.hwndRender) return;

    EnsureGuardClass(GetModuleHandleW(nullptr));

    UINT dpi = GetDpiForWindow(g_state.hwndRender);
    if (dpi == 0) dpi = USER_DEFAULT_SCREEN_DPI;
    // Cover the full width of the render pane and the top ~40 px where the
    // Office ribbon / title bar lives. This blocks both the close button
    // and the context menu that appears on a right-click in the title area.
    RECT rc;
    GetClientRect(g_state.hwndRender, &rc);
    int w = rc.right - rc.left;
    int h = ScaleForDpi(40, dpi);
    int x = 0;
    int y = 0;

    g_state.hwndCloseGuard = CreateWindowExW(
        0,
        kGuardClassName, L"",
        WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
        x, y, w, h,
        g_state.hwndRender, nullptr, GetModuleHandleW(nullptr), nullptr);

    if (g_state.hwndCloseGuard)
    {
        // Force the guard to the top of the Z-order so it sits above the
        // Office window that was just reparented into hwndRender.
        SetWindowPos(g_state.hwndCloseGuard, HWND_TOP,
                     0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        HostLog(L"CreateCloseGuard: created at (%d,%d) size=%dx%d", x, y, w, h);
    }
    else
    {
        HostLog(L"CreateCloseGuard: CreateWindowEx failed err=%lu", GetLastError());
    }
}

static void DestroyCloseGuard()
{
    if (g_state.hwndCloseGuard)
    {
        DestroyWindow(g_state.hwndCloseGuard);
        g_state.hwndCloseGuard = nullptr;
    }
}

// Reparent an Office app's main HWND into our render window and strip its
// decorations so it looks like an embedded preview pane instead of a
// top-level frame. Caller stores the HWND in the appropriate HostState slot.
static bool EmbedOfficeWindowSta(HWND hwndApp)
{
    // Hide the window before any style/parent changes. Cross-process
    // SetParent on a visible top-level window forces an expensive
    // synchronous redraw and focus dance across process boundaries.
    // PowerPoint in particular refuses Visible=False, so its frame is
    // on-screen from the moment CoCreateInstance returns; hiding it
    // here makes the reparent cheap and avoids the flash.
    ShowWindow(hwndApp, SW_HIDE);

    LONG_PTR style   = GetWindowLongPtrW(hwndApp, GWL_STYLE);
    LONG_PTR exStyle = GetWindowLongPtrW(hwndApp, GWL_EXSTYLE);

    style &= ~(WS_OVERLAPPED | WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX |
               WS_MAXIMIZEBOX | WS_SYSMENU | WS_DLGFRAME | WS_BORDER | WS_POPUP);
    style |= WS_CHILD | WS_CLIPCHILDREN;
    exStyle &= ~(WS_EX_DLGMODALFRAME | WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE |
                 WS_EX_STATICEDGE   | WS_EX_APPWINDOW   | WS_EX_TOOLWINDOW);
    SetWindowLongPtrW(hwndApp, GWL_STYLE,   style);
    SetWindowLongPtrW(hwndApp, GWL_EXSTYLE, exStyle);

    if (!SetParent(hwndApp, g_state.hwndRender))
    {
        HostLog(L"  EmbedOfficeWindowSta: SetParent failed err=%lu", GetLastError());
        return false;
    }

    // Create an invisible overlay that covers the Office app's own close
    // button in the top-right corner. Office runs out-of-process, so
    // SetWindowSubclass does not work on its HWND; the guard window (a
    // child of our render pane in this process) is the only reliable way
    // to block clicks on the button.
    CreateCloseGuard();

    RECT rc; GetClientRect(g_state.hwndRender, &rc);
    SetWindowPos(hwndApp, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top,
                 SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_SHOWWINDOW);
    return true;
}

// Add an Office process (Word / Excel / PowerPoint) to our kill-on-close
// Job Object, so that whenever this host dies the Office instance dies
// with it. Idempotent: re-assigning a process that is already a member
// of the job is a no-op (AssignProcessToJobObject returns FALSE but
// that's fine — we just log and move on). On modern Windows (8+) nested
// jobs work transparently, so even if the Office process was placed in
// some other job by the shell, our job's KILL_ON_JOB_CLOSE still fires.
static void AssignOfficeProcessToJobSta(HWND hwndOfficeApp)
{
    if (!g_state.hOfficeJob || !hwndOfficeApp) return;
    DWORD pid = 0;
    GetWindowThreadProcessId(hwndOfficeApp, &pid);
    if (!pid) return;

    HANDLE hProc = OpenProcess(PROCESS_TERMINATE | PROCESS_SET_QUOTA,
                               FALSE, pid);
    if (!hProc)
    {
        HostLog(L"  AssignOfficeProcessToJob: OpenProcess(pid=%lu) failed err=%lu",
                pid, GetLastError());
        return;
    }
    BOOL ok = AssignProcessToJobObject(g_state.hOfficeJob, hProc);
    HostLog(L"  AssignProcessToJobObject(pid=%lu) -> %s (err=%lu)",
            pid, ok ? L"OK" : L"FAIL",
            ok ? 0UL : GetLastError());
    CloseHandle(hProc);
}

static HRESULT LoadWordFullSta(LPCWSTR path)
{
    HostLog(L"LoadWordFullSta: path='%s'", path);

    // Only one full-mode app can be embedded in the render pane at a time.
    // Tear down any leftover quick-mode handler and any other Office app
    // before installing Word.
    UnloadHandlerSta();
    UnloadExcelFullSta(true);
    UnloadPptFullSta(true);

    // (Re)use the live Word.Application if we already created one in this
    // session; otherwise spin one up.
    if (!g_state.pWordApp)
    {
        CLSID clsid;
        HRESULT hr = CLSIDFromProgID(L"Word.Application", &clsid);
        HostLog(L"  CLSIDFromProgID(Word.Application) -> 0x%08lX", static_cast<long>(hr));
        if (FAILED(hr)) return hr;

        IDispatch* pApp = nullptr;
        hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER,
                              IID_IDispatch, reinterpret_cast<void**>(&pApp));
        HostLog(L"  CoCreateInstance(Word.Application) -> 0x%08lX", static_cast<long>(hr));
        if (FAILED(hr)) return hr;

        // Suppress modal dialogs: wdAlertsNone (= 0).
        VARIANT vAlerts; VariantInit(&vAlerts);
        vAlerts.vt = VT_I4; vAlerts.lVal = 0;
        DispCall(pApp, L"DisplayAlerts", DISPATCH_PROPERTYPUT, &vAlerts, 1, nullptr);
        // Keep Word logically invisible until we've reparented its window.
        DispPutBool(pApp, L"Visible", false);

        g_state.pWordApp = pApp;
    }
    else
    {
        // Reusing — make sure no stale doc and no stale embedding remain.
        CloseWordDocumentSta();
        DetachWordWindowSta();
    }

    // Word's Documents.Open signature is (FileName, ConfirmConversions,
    // ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert,
    // WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible,
    // ...). Skipping parameters in positional IDispatch::Invoke is not
    // possible (you'd need named args or VT_ERROR/DISP_E_PARAMNOTFOUND), so
    // we pass exactly the first four positional params and rely on
    // Application.Visible=False (set just above) to keep the document hidden
    // until reparenting completes. IDispatch::Invoke takes args in REVERSE
    // positional order.
    VARIANT vDocs; VariantInit(&vDocs);
    HRESULT hr = DispGetProperty(g_state.pWordApp, L"Documents", &vDocs);
    if (FAILED(hr) || vDocs.vt != VT_DISPATCH)
    {
        HostLog(L"  get Documents -> 0x%08lX vt=%d", static_cast<long>(hr), vDocs.vt);
        VariantClear(&vDocs);
        return FAILED(hr) ? hr : E_FAIL;
    }
    IDispatch* pDocs = vDocs.pdispVal;

    // Raw long path — Office rejects the \\?\ prefix on BSTR FileName
    // arguments the same way its preview handlers do.  Long paths work
    // here only because Word ships with longPathAware itself and the
    // system has LongPathsEnabled=1 in the registry.  Excel and
    // PowerPoint have their own internal ~218-char limit on this
    // argument that no system setting can lift.
    VARIANT args[4];
    for (int i = 0; i < 4; ++i) VariantInit(&args[i]);
    args[0].vt = VT_BOOL; args[0].boolVal = VARIANT_FALSE;            // AddToRecentFiles
    args[1].vt = VT_BOOL; args[1].boolVal = VARIANT_TRUE;             // ReadOnly
    args[2].vt = VT_BOOL; args[2].boolVal = VARIANT_FALSE;            // ConfirmConversions
    args[3].vt = VT_BSTR; args[3].bstrVal = SysAllocString(path);     // FileName

    VARIANT vDoc; VariantInit(&vDoc);
    hr = DispCall(pDocs, L"Open", DISPATCH_METHOD, args, 4, &vDoc);
    HostLog(L"  Documents.Open -> 0x%08lX vt=%d", static_cast<long>(hr), vDoc.vt);
    pDocs->Release();
    SysFreeString(args[3].bstrVal);

    if (FAILED(hr) || vDoc.vt != VT_DISPATCH)
    {
        VariantClear(&vDoc);
        return FAILED(hr) ? hr : E_FAIL;
    }
    g_state.pWordDoc = vDoc.pdispVal;       // owned

    // Word does not materialise Application.Hwnd until the app is made
    // visible. There's no way around a brief on-screen flash: we set
    // Visible=True, read the HWND, then immediately SetParent it into
    // our render pane. Empirically the flash is short enough that users
    // perceive it as "Word's icon blipped on the taskbar".
    DispPutBool(g_state.pWordApp, L"Visible", true);

    // Modern Microsoft 365 Word does not expose Application.Hwnd via
    // IDispatch (GetIDsOfNames returns DISP_E_UNKNOWNNAME), so we try once
    // for backwards compatibility with older Office and immediately fall
    // back to a class-name scan otherwise.
    HWND hwndWord = DispGetHwndProperty(g_state.pWordApp, L"Hwnd");
    if (!hwndWord)
    {
        LPCWSTR fname = wcsrchr(path, L'\\');
        fname = fname ? fname + 1 : path;
        // Word can take a moment to create its top-level window after
        // Visible=True; retry the EnumWindows scan a few times.
        // We deliberately do NOT pump messages here — PeekMessage would
        // dispatch WM_HOST_SWITCH_MODE and recurse into LoadFileWithModeSta
        // while we are still mid-load, corrupting COM state.
        for (int retry = 0; retry < 40 && !hwndWord; ++retry)
        {
            hwndWord = FindOfficeTopLevelWindow(L"OpusApp", fname);
            if (hwndWord) break;
            Sleep(25);
        }
        HostLog(L"  FindOfficeTopLevelWindow(OpusApp '%s') = 0x%p",
                fname, hwndWord);
        if (!hwndWord)
        {
            hwndWord = FindOfficeTopLevelWindow(L"OpusApp", nullptr);
            HostLog(L"  FindOfficeTopLevelWindow(OpusApp any) = 0x%p", hwndWord);
        }
    }
    if (!hwndWord)
    {
        CloseWordDocumentSta();
        return E_FAIL;
    }
    if (!EmbedOfficeWindowSta(hwndWord))
    {
        CloseWordDocumentSta();
        return E_FAIL;
    }
    g_state.hwndWordApp = hwndWord;
    AssignOfficeProcessToJobSta(hwndWord);

    // Apply preview-friendly tweaks: Print Layout view, read-only
    // protection, hidden status bar / rulers.
    ConfigureWordForPreviewSta();

    HostLog(L"  LoadWordFullSta SUCCESS");
    return S_OK;
}

// ---------------------------------------------------------------------------
// Full-mode Excel lifecycle. Mirrors the Word path but with Excel-specific
// COM names: Excel.Application / Workbooks / Workbook. Excel can run with
// Visible=False until embedded, so the brief on-screen flash before
// reparenting is short.
// ---------------------------------------------------------------------------

static void CloseExcelWorkbookSta()
{
    if (!g_state.pExcelWb) return;
    // Excel.Workbook.Close(SaveChanges) — Boolean (not the Word enum).
    VARIANT vSave; VariantInit(&vSave);
    vSave.vt = VT_BOOL; vSave.boolVal = VARIANT_FALSE;
    DispCall(g_state.pExcelWb, L"Close", DISPATCH_METHOD, &vSave, 1, nullptr);
    g_state.pExcelWb->Release();
    g_state.pExcelWb = nullptr;
}

static void DetachExcelWindowSta()
{
    if (!g_state.hwndExcelApp) return;
    ShowWindow(g_state.hwndExcelApp, SW_HIDE);
    DestroyCloseGuard();
    SetParent(g_state.hwndExcelApp, nullptr);
    LONG_PTR style = GetWindowLongPtrW(g_state.hwndExcelApp, GWL_STYLE);
    style &= ~WS_CHILD;
    style |= WS_OVERLAPPEDWINDOW;
    SetWindowLongPtrW(g_state.hwndExcelApp, GWL_STYLE, style);
    SetWindowPos(g_state.hwndExcelApp, nullptr, 0, 0, 0, 0,
                 SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE |
                 SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);
    g_state.hwndExcelApp = nullptr;
}

static void UnloadExcelFullSta(bool quitApp)
{
    CloseExcelWorkbookSta();
    DetachExcelWindowSta();
    if (quitApp && g_state.pExcelApp)
    {
        DispCall(g_state.pExcelApp, L"Quit", DISPATCH_METHOD, nullptr, 0, nullptr);
        g_state.pExcelApp->Release();
        g_state.pExcelApp = nullptr;
    }
}

static HRESULT LoadExcelFullSta(LPCWSTR path)
{
    HostLog(L"LoadExcelFullSta: path='%s'", path);

    // Only one full-mode app at a time; tear down the others.
    UnloadHandlerSta();
    UnloadWordFullSta(true);
    UnloadPptFullSta(true);                  // forward-declared below

    if (!g_state.pExcelApp)
    {
        CLSID clsid;
        HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
        HostLog(L"  CLSIDFromProgID(Excel.Application) -> 0x%08lX",
                static_cast<long>(hr));
        if (FAILED(hr)) return hr;

        IDispatch* pApp = nullptr;
        hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER,
                              IID_IDispatch, reinterpret_cast<void**>(&pApp));
        HostLog(L"  CoCreateInstance(Excel.Application) -> 0x%08lX",
                static_cast<long>(hr));
        if (FAILED(hr)) return hr;

        // Excel.DisplayAlerts is a Boolean.
        DispPutBool(pApp, L"DisplayAlerts", false);
        DispPutBool(pApp, L"Visible", false);
        g_state.pExcelApp = pApp;
    }
    else
    {
        CloseExcelWorkbookSta();
        DetachExcelWindowSta();
    }

    // Pre-size Excel to match our render pane *before* opening the
    // workbook. Excel persists its frame size in the user profile, so
    // a freshly created instance starts at whatever size the user's
    // last standalone Excel session was — usually wildly different from
    // our Lister pane. If we let Excel lay the workbook out at that
    // saved size and then resize the Win32 frame, Excel's child widgets
    // (ribbon, sheet tabs, status bar) stay anchored to the old size
    // and end up clipped or stranded. Setting Width/Height up-front
    // makes Excel's initial layout match the embed target.
    {
        RECT rc; GetClientRect(g_state.hwndRender, &rc);
        ApplyExcelAppSizeSta(rc.right - rc.left, rc.bottom - rc.top);
    }

    // Workbooks.Open(FileName, UpdateLinks, ReadOnly, ...). Args in REVERSE
    // positional order. UpdateLinks=0 (xlUpdateLinksNever) prevents the
    // "do you want to update links?" dialog from blocking us.
    VARIANT vWbs; VariantInit(&vWbs);
    HRESULT hr = DispGetProperty(g_state.pExcelApp, L"Workbooks", &vWbs);
    if (FAILED(hr) || vWbs.vt != VT_DISPATCH)
    {
        HostLog(L"  get Workbooks -> 0x%08lX vt=%d",
                static_cast<long>(hr), vWbs.vt);
        VariantClear(&vWbs);
        return FAILED(hr) ? hr : E_FAIL;
    }
    IDispatch* pWbs = vWbs.pdispVal;

    // Raw long path; Office rejects \\?\ on BSTR FileName.  Excel has
    // its own ~218-char internal limit beyond what the system permits.
    VARIANT args[3];
    for (int i = 0; i < 3; ++i) VariantInit(&args[i]);
    args[0].vt = VT_BOOL; args[0].boolVal = VARIANT_TRUE;             // ReadOnly
    args[1].vt = VT_I4;   args[1].lVal    = 0;                        // UpdateLinks = xlUpdateLinksNever
    args[2].vt = VT_BSTR; args[2].bstrVal = SysAllocString(path);     // FileName

    VARIANT vWb; VariantInit(&vWb);
    hr = DispCall(pWbs, L"Open", DISPATCH_METHOD, args, 3, &vWb);
    HostLog(L"  Workbooks.Open -> 0x%08lX vt=%d",
            static_cast<long>(hr), vWb.vt);
    pWbs->Release();
    SysFreeString(args[2].bstrVal);

    if (FAILED(hr) || vWb.vt != VT_DISPATCH)
    {
        VariantClear(&vWb);
        return FAILED(hr) ? hr : E_FAIL;
    }
    g_state.pExcelWb = vWb.pdispVal;

    DispPutBool(g_state.pExcelApp, L"Visible", true);

    HWND hwndExcel = DispGetHwndProperty(g_state.pExcelApp, L"Hwnd");
    if (!hwndExcel)
    {
        LPCWSTR fname = wcsrchr(path, L'\\');
        fname = fname ? fname + 1 : path;
        for (int retry = 0; retry < 40 && !hwndExcel; ++retry)
        {
            hwndExcel = FindOfficeTopLevelWindow(L"XLMAIN", fname);
            if (hwndExcel) break;
            Sleep(25);
        }
        HostLog(L"  FindOfficeTopLevelWindow(XLMAIN '%s') = 0x%p",
                fname, hwndExcel);
        if (!hwndExcel)
        {
            hwndExcel = FindOfficeTopLevelWindow(L"XLMAIN", nullptr);
            HostLog(L"  FindOfficeTopLevelWindow(XLMAIN any) = 0x%p", hwndExcel);
        }
    }
    if (!hwndExcel)
    {
        CloseExcelWorkbookSta();
        return E_FAIL;
    }
    if (!EmbedOfficeWindowSta(hwndExcel))
    {
        CloseExcelWorkbookSta();
        return E_FAIL;
    }
    g_state.hwndExcelApp = hwndExcel;
    AssignOfficeProcessToJobSta(hwndExcel);

    ConfigureExcelForPreviewSta();

    HostLog(L"  LoadExcelFullSta SUCCESS");
    return S_OK;
}

// ---------------------------------------------------------------------------
// Full-mode PowerPoint lifecycle. Unlike Word and Excel, PowerPoint cannot
// run with Application.Visible = False — the property rejects msoFalse —
// so its main window is always on-screen between CoCreateInstance and our
// reparent. The visible flash is brief but unavoidable.
// ---------------------------------------------------------------------------

static void ClosePptPresentationSta()
{
    if (!g_state.pPptPres) return;
    // Tell PowerPoint there are no unsaved changes so Close() doesn't prompt.
    DispPutBool(g_state.pPptPres, L"Saved", true);
    DispCall(g_state.pPptPres, L"Close", DISPATCH_METHOD, nullptr, 0, nullptr);
    g_state.pPptPres->Release();
    g_state.pPptPres = nullptr;
}

static void DetachPptWindowSta()
{
    if (!g_state.hwndPptApp) return;
    ShowWindow(g_state.hwndPptApp, SW_HIDE);
    DestroyCloseGuard();
    SetParent(g_state.hwndPptApp, nullptr);
    LONG_PTR style = GetWindowLongPtrW(g_state.hwndPptApp, GWL_STYLE);
    style &= ~WS_CHILD;
    style |= WS_OVERLAPPEDWINDOW;
    SetWindowLongPtrW(g_state.hwndPptApp, GWL_STYLE, style);
    SetWindowPos(g_state.hwndPptApp, nullptr, 0, 0, 0, 0,
                 SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE |
                 SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);
    g_state.hwndPptApp = nullptr;
}

static void UnloadPptFullSta(bool quitApp)
{
    ClosePptPresentationSta();
    DetachPptWindowSta();
    if (quitApp && g_state.pPptApp)
    {
        if (!g_state.pPptAppIsShared)
        {
            // Only quit when we own the process; quitting a shared (pre-existing)
            // instance would close all the user's open presentations.
            DispCall(g_state.pPptApp, L"Quit", DISPATCH_METHOD, nullptr, 0, nullptr);
        }
        g_state.pPptApp->Release();
        g_state.pPptApp = nullptr;
        g_state.pPptAppIsShared = false;
    }
}

static HRESULT LoadPowerPointFullSta(LPCWSTR path)
{
    HostLog(L"LoadPowerPointFullSta: path='%s'", path);

    UnloadHandlerSta();
    UnloadWordFullSta(true);
    UnloadExcelFullSta(true);

    if (!g_state.pPptApp)
    {
        // Detect whether a PowerPoint process is already running before we
        // call CoCreateInstance.  PowerPoint registers with REGCLS_MULTIPLEUSE,
        // so if it is already running CoCreateInstance returns a proxy to the
        // existing process rather than starting a new one.  In that case we
        // must not call Application.Quit on unload (it would close all the
        // user's open presentations) and must not assign the process to our
        // kill-on-close job object.
        const bool pptWasRunning =
            (FindOfficeTopLevelWindow(L"PPTFrameClass", nullptr) != nullptr);
        HostLog(L"  PowerPoint pre-existing instance: %s",
                pptWasRunning ? L"yes (shared — will not Quit or assign to job)"
                              : L"no (new process)");

        CLSID clsid;
        HRESULT hr = CLSIDFromProgID(L"PowerPoint.Application", &clsid);
        HostLog(L"  CLSIDFromProgID(PowerPoint.Application) -> 0x%08lX",
                static_cast<long>(hr));
        if (FAILED(hr)) return hr;

        IDispatch* pApp = nullptr;
        hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER,
                              IID_IDispatch, reinterpret_cast<void**>(&pApp));
        HostLog(L"  CoCreateInstance(PowerPoint.Application) -> 0x%08lX",
                static_cast<long>(hr));
        if (FAILED(hr)) return hr;

        // PowerPoint.DisplayAlerts uses an enum (ppAlertsNone = 1).
        VARIANT vAlerts; VariantInit(&vAlerts);
        vAlerts.vt = VT_I4; vAlerts.lVal = 1;
        DispCall(pApp, L"DisplayAlerts", DISPATCH_PROPERTYPUT, &vAlerts, 1, nullptr);
        // PowerPoint refuses Visible=False. The window will be on screen
        // between Presentations.Open and our reparent — accept the brief flash.

        g_state.pPptApp = pApp;
        g_state.pPptAppIsShared = pptWasRunning;
    }
    else
    {
        ClosePptPresentationSta();
        DetachPptWindowSta();
    }

    // Presentations.Open(FileName, ReadOnly, Untitled, WithWindow). Args
    // in REVERSE positional order. Microsoft constants are MsoTriState:
    // msoTrue = -1, msoFalse = 0.
    VARIANT vPress; VariantInit(&vPress);
    HRESULT hr = DispGetProperty(g_state.pPptApp, L"Presentations", &vPress);
    if (FAILED(hr) || vPress.vt != VT_DISPATCH)
    {
        HostLog(L"  get Presentations -> 0x%08lX vt=%d",
                static_cast<long>(hr), vPress.vt);
        VariantClear(&vPress);
        return FAILED(hr) ? hr : E_FAIL;
    }
    IDispatch* pPresentations = vPress.pdispVal;

    // Raw long path; Office rejects \\?\ on BSTR FileName.  PowerPoint
    // has its own internal limit beyond what the system permits.
    VARIANT args[4];
    for (int i = 0; i < 4; ++i) VariantInit(&args[i]);
    args[0].vt = VT_I4;   args[0].lVal    = -1;                       // WithWindow = msoTrue
    args[1].vt = VT_I4;   args[1].lVal    = 0;                        // Untitled = msoFalse
    args[2].vt = VT_I4;   args[2].lVal    = -1;                       // ReadOnly = msoTrue
    args[3].vt = VT_BSTR; args[3].bstrVal = SysAllocString(path);     // FileName

    VARIANT vPres; VariantInit(&vPres);
    hr = DispCall(pPresentations, L"Open", DISPATCH_METHOD, args, 4, &vPres);
    HostLog(L"  Presentations.Open -> 0x%08lX vt=%d",
            static_cast<long>(hr), vPres.vt);
    pPresentations->Release();
    SysFreeString(args[3].bstrVal);

    if (FAILED(hr) || vPres.vt != VT_DISPATCH)
    {
        VariantClear(&vPres);
        return FAILED(hr) ? hr : E_FAIL;
    }
    g_state.pPptPres = vPres.pdispVal;

    HWND hwndPpt = DispGetHwndProperty(g_state.pPptApp, L"HWND");        // PowerPoint uses uppercase HWND
    if (!hwndPpt)
        hwndPpt = DispGetHwndProperty(g_state.pPptApp, L"Hwnd");         // fall back to mixed case
    if (!hwndPpt)
    {
        LPCWSTR fname = wcsrchr(path, L'\\');
        fname = fname ? fname + 1 : path;
        for (int retry = 0; retry < 40 && !hwndPpt; ++retry)
        {
            hwndPpt = FindOfficeTopLevelWindow(L"PPTFrameClass", fname);
            if (hwndPpt) break;
            Sleep(25);
        }
        HostLog(L"  FindOfficeTopLevelWindow(PPTFrameClass '%s') = 0x%p",
                fname, hwndPpt);
        if (!hwndPpt)
        {
            hwndPpt = FindOfficeTopLevelWindow(L"PPTFrameClass", nullptr);
            HostLog(L"  FindOfficeTopLevelWindow(PPTFrameClass any) = 0x%p",
                    hwndPpt);
        }
    }
    if (!hwndPpt)
    {
        ClosePptPresentationSta();
        return E_FAIL;
    }
    if (!EmbedOfficeWindowSta(hwndPpt))
    {
        ClosePptPresentationSta();
        return E_FAIL;
    }
    g_state.hwndPptApp = hwndPpt;
    // Only assign to the kill-on-close job when we own the process.
    // For a shared (pre-existing) instance assigning it would kill all
    // the user's open presentations when the host exits.
    if (!g_state.pPptAppIsShared)
        AssignOfficeProcessToJobSta(hwndPpt);

    ConfigurePowerPointForPreviewSta();

    HostLog(L"  LoadPowerPointFullSta SUCCESS");
    return S_OK;
}

static void ResizeOfficeFullSta(int w, int h)
{
    const UINT kFlags = SWP_NOZORDER | SWP_NOACTIVATE;

    // Only one of the three is non-null at any moment, but checking all
    // is cheap and avoids having to track "currently active app" state.
    if (g_state.hwndWordApp)
    {
        SetWindowPos(g_state.hwndWordApp, nullptr, 0, 0, w, h, kFlags);
    }
    if (g_state.hwndExcelApp)
    {
        // Plain Win32 resize. Excel won't actually relayout its child
        // widgets on this (see comment in ConfigureExcelForPreviewSta);
        // we set the frame size for consistency with Word / PowerPoint
        // even though the visible content stays at its load-time layout.
        SetWindowPos(g_state.hwndExcelApp, nullptr, 0, 0, w, h, kFlags);
    }
    if (g_state.hwndPptApp)
    {
        SetWindowPos(g_state.hwndPptApp, nullptr, 0, 0, w, h, kFlags);
    }
}

// ---------------------------------------------------------------------------
// Preview-handler site object.
//
// IPreviewHandler requires the host to supply a site that implements:
//
//   IPreviewHandlerFrame  — accelerator translation contract; required by
//                           Outlook's MSG handler and several other modern
//                           Office handlers.  Without it DoPreview()
//                           returns E_FAIL.
//   IServiceProvider      — routes service queries (handlers often look up
//                           IPreviewHandlerFrame through this).
//   IOleWindow            — gives the handler a stable parent HWND it can
//                           query at any time without depending on the
//                           SetWindow argument staying valid.  Outlook in
//                           particular needs this during DoPreview.
//
// The site must be installed via IObjectWithSite::SetSite BEFORE the
// initial Initialize(*) call: Outlook's MSG handler asks for the site
// from inside Initialize and aborts the load if it isn't there yet.
// ---------------------------------------------------------------------------

class PreviewHostSite : public IPreviewHandlerFrame,
                        public IServiceProvider,
                        public IOleWindow
{
    HWND m_hwndHost;
    LONG m_ref = 1;
public:
    explicit PreviewHostSite(HWND hwndHost) : m_hwndHost(hwndHost) {}

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == __uuidof(IPreviewHandlerFrame))
            *ppv = static_cast<IPreviewHandlerFrame*>(this);
        else if (riid == IID_IServiceProvider)
            *ppv = static_cast<IServiceProvider*>(this);
        else if (riid == __uuidof(IOleWindow))
            *ppv = static_cast<IOleWindow*>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    STDMETHODIMP_(ULONG) AddRef() override
    {
        return InterlockedIncrement(&m_ref);
    }
    STDMETHODIMP_(ULONG) Release() override
    {
        LONG r = InterlockedDecrement(&m_ref);
        if (r == 0) delete this;
        return r;
    }

    // IPreviewHandlerFrame
    STDMETHODIMP GetWindowContext(PREVIEWHANDLERFRAMEINFO* pInfo) override
    {
        if (!pInfo) return E_POINTER;
        pInfo->haccel        = nullptr;
        pInfo->cAccelEntries = 0;
        return S_OK;
    }
    STDMETHODIMP TranslateAccelerator(MSG* /*pMsg*/) override
    {
        return S_FALSE;          // we don't consume any accelerator
    }

    // IServiceProvider — route IPreviewHandlerFrame and IOleWindow lookups
    // back to ourselves (handlers sometimes ask through this rather than QI).
    STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (guidService == __uuidof(IPreviewHandlerFrame) ||
            guidService == __uuidof(IOleWindow))
            return QueryInterface(riid, ppv);
        return E_NOINTERFACE;
    }

    // IOleWindow — return our render window so handlers can compute layout,
    // create child controls, etc. independently of the SetWindow argument.
    STDMETHODIMP GetWindow(HWND* phwnd) override
    {
        if (!phwnd) return E_POINTER;
        *phwnd = m_hwndHost;
        return S_OK;
    }
    STDMETHODIMP ContextSensitiveHelp(BOOL /*fEnter*/) override
    {
        return E_NOTIMPL;
    }
};

// ---------------------------------------------------------------------------
// Preview handler lifecycle. STA thread only.
// ---------------------------------------------------------------------------

// Detach any child window under hwndRender that isn't one of ours.
//
// Preview handlers run in the prevhost.exe surrogate and reparent their own
// rendering window under our hwndRender via cross-process SetParent.  When we
// release the handler (IPreviewHandler::Unload + Release) the surrogate is
// SUPPOSED to destroy that window, but in practice — especially for handlers
// that crashed mid-DoPreview, were swapped out without a clean shutdown, or
// hold references on a background thread — the window can linger.  A leftover
// child sits on top of whatever we paint next (mode/unblock buttons, fallback
// EDIT) and shows the previous file's content through the gaps.
//
// We can't destroy a window owned by another process, but we CAN SetParent it
// elsewhere — Win32 permits cross-process parent changes.  Moving the orphan
// to HWND_MESSAGE takes it out of hwndRender's visible hierarchy entirely;
// the surrogate retains ownership and will dispose of it on its own schedule.
static void PurgeOrphanRenderChildrenSta()
{
    if (!g_state.hwndRender) return;

    // Snapshot first, then act — SetParent mutates the sibling chain we'd be
    // walking with GetWindow(GW_HWNDNEXT).
    std::vector<HWND> orphans;
    HWND child = GetWindow(g_state.hwndRender, GW_CHILD);
    while (child)
    {
        if (child != g_state.hwndModeButton    &&
            child != g_state.hwndUnblockButton &&
            child != g_state.hwndFallback      &&
            child != g_state.hwndCloseGuard    &&
            child != g_state.hwndWordApp       &&
            child != g_state.hwndExcelApp      &&
            child != g_state.hwndPptApp)
        {
            orphans.push_back(child);
        }
        child = GetWindow(child, GW_HWNDNEXT);
    }

    for (HWND w : orphans)
    {
        // SetParent to HWND_MESSAGE detaches w from the visible hierarchy
        // without destroying it.  Best-effort: failure (e.g. the surrogate
        // already tore the window down between enumeration and call) is
        // benign — log and move on.
        HWND prev = SetParent(w, HWND_MESSAGE);
        HostLog(L"  PurgeOrphanRenderChildrenSta: detached hwnd=0x%p (was child of 0x%p)",
                w, prev);
    }
}

static void UnloadHandlerSta()
{
    HideFallbackSta();
    if (g_state.pHandler)
    {
        // Drop our site reference cleanly so the handler can release its
        // back-pointer to it before we tear the handler down.
        IObjectWithSite* pOws = nullptr;
        if (SUCCEEDED(g_state.pHandler->QueryInterface(
                __uuidof(IObjectWithSite), reinterpret_cast<void**>(&pOws))))
        {
            pOws->SetSite(nullptr);
            pOws->Release();
        }
        g_state.pHandler->Unload();
        g_state.pHandler->Release();
        g_state.pHandler = nullptr;
    }
    if (g_state.pHandlerUnk)
    {
        g_state.pHandlerUnk->Release();
        g_state.pHandlerUnk = nullptr;
    }
    // After COM has released its references, sweep up any cross-process
    // child window the surrogate left behind under hwndRender.  Must run
    // AFTER the IPreviewHandler release so we don't yank a window that
    // the handler is still drawing into.
    PurgeOrphanRenderChildrenSta();
}

// Attempt a quick-mode preview-handler load.  Returns S_OK on real success
// (g_state.pHandler / pHandlerUnk now hold live references) or the actual
// failure HRESULT.  Does NOT show the fallback panel — that's the caller's
// decision, so quick-mode failure can be retried in full mode for Office
// files before the user ever sees an error.
static HRESULT TryLoadHandlerSta(LPCWSTR path)
{
    UnloadHandlerSta();
    HostLog(L"TryLoadHandlerSta: path='%s' render=0x%p", path, g_state.hwndRender);

    // Long-path policy in this function: pass the RAW path everywhere.
    //
    // We learned the hard way that none of the consumers below accept the
    // \\?\ prefix:
    //   - Shell APIs (SHCreateItemFromParsingName / SHCreateStreamOnFileEx)
    //     reject \\?\ in "display names" with E_INVALIDARG.  They handle
    //     long paths through the system's longPathAware manifest + the
    //     HKLM\…\FileSystem\LongPathsEnabled registry switch instead.
    //   - Office preview handlers (Word / Excel / PowerPoint) likewise
    //     reject \\?\ with E_NOTIMPL on IInitializeWithFile::Initialize.
    //
    // EnsureLongPathPrefix is still useful for Win32 file APIs that DO
    // accept \\?\ (notably GetFileAttributesExW in BuildFallbackText), so
    // the helper stays.  It just isn't applied here.

    CLSID clsid;
    HRESULT hr = FindPreviewHandlerClsid(path, &clsid);
    if (FAILED(hr))
    {
        HostLog(L"  FindPreviewHandlerClsid -> 0x%08lX",
                static_cast<long>(hr));
        return hr;
    }
    {
        wchar_t clsidStr[64] = {};
        StringFromGUID2(clsid, clsidStr, ARRAYSIZE(clsidStr));
        HostLog(L"  CLSID = %s", clsidStr);
    }

    // CLSCTX_LOCAL_SERVER only (no CLSCTX_INPROC_SERVER): preview handlers
    // are contractually required to run out-of-process — Microsoft's
    // IPreviewHandler docs state "A preview handler must be hosted in its
    // own background process" — and many handlers actively check for this.
    // Windows' built-in MAPI Mail Previewer (mssvp.dll, the .msg handler),
    // for example, succeeds CoCreateInstance / Initialize / SetWindow but
    // then fails DoPreview with E_FAIL when loaded in-process even though
    // its registration has both InProcServer32 and an AppID surrogate.
    // Forcing LOCAL_SERVER routes activation through the AppID's
    // DllSurrogate (typically prevhost.exe), matching Explorer's preview
    // pane and Microsoft's reference preview-host sample.
    IUnknown* pUnk = nullptr;
    hr = CoCreateInstance(clsid, nullptr,
                          CLSCTX_LOCAL_SERVER,
                          IID_IUnknown, reinterpret_cast<void**>(&pUnk));
    HostLog(L"  CoCreateInstance -> 0x%08lX", static_cast<long>(hr));
    if (FAILED(hr))
    {
        // Handler is registered but the server cannot be activated — typically
        // a partial / broken Office install.  Caller decides whether to retry
        // in full mode (for Office files) or show the fallback panel.
        return hr;
    }

    // Microsoft's recommended call order for hosting a preview handler is:
    //   CoCreateInstance → QI IPreviewHandler → SetSite → Initialize* →
    //   SetWindow → DoPreview.
    // Outlook's MSG handler in particular asks for its site from inside
    // Initialize() and fails the whole load with E_FAIL if it isn't there
    // yet, so the SetSite block has to run BEFORE any Initialize* call.

    IPreviewHandler* pPH = nullptr;
    hr = pUnk->QueryInterface(IID_IPreviewHandler, reinterpret_cast<void**>(&pPH));
    HostLog(L"  QI IPreviewHandler -> 0x%08lX", static_cast<long>(hr));
    if (FAILED(hr))
    {
        pUnk->Release();
        return hr;
    }

    // Install the site (best-effort: a handler that doesn't implement
    // IObjectWithSite simply doesn't need one).
    {
        IObjectWithSite* pOws = nullptr;
        HRESULT shr = pPH->QueryInterface(__uuidof(IObjectWithSite),
                                          reinterpret_cast<void**>(&pOws));
        HostLog(L"  QI IObjectWithSite -> 0x%08lX", static_cast<long>(shr));
        if (SUCCEEDED(shr))
        {
            PreviewHostSite* pSite = new PreviewHostSite(g_state.hwndRender);
            HRESULT ssr = pOws->SetSite(static_cast<IPreviewHandlerFrame*>(pSite));
            HostLog(L"  IObjectWithSite::SetSite -> 0x%08lX", static_cast<long>(ssr));
            pSite->Release();              // SetSite holds its own ref if it needs one
            pOws->Release();
        }
    }

    bool initialized = false;

    {
        IInitializeWithFile* pInit = nullptr;
        HRESULT qhr = pUnk->QueryInterface(__uuidof(IInitializeWithFile),
                                           reinterpret_cast<void**>(&pInit));
        HostLog(L"  QI IInitializeWithFile -> 0x%08lX", static_cast<long>(qhr));
        if (SUCCEEDED(qhr))
        {
            // Raw path, not the \\?\-prefixed form: Office's preview
            // handlers actively reject the prefix with E_NOTIMPL, but
            // accept raw long paths when the system has LongPathsEnabled.
            HRESULT ihr = pInit->Initialize(path, STGM_READ);
            HostLog(L"  IInitializeWithFile::Initialize -> 0x%08lX",
                    static_cast<long>(ihr));
            pInit->Release();
            initialized = SUCCEEDED(ihr);
            if (!initialized) hr = ihr;
        }
    }

    if (!initialized)
    {
        IInitializeWithStream* pInit = nullptr;
        HRESULT qhr = pUnk->QueryInterface(__uuidof(IInitializeWithStream),
                                           reinterpret_cast<void**>(&pInit));
        HostLog(L"  QI IInitializeWithStream -> 0x%08lX", static_cast<long>(qhr));
        if (SUCCEEDED(qhr))
        {
            IStream* pStream = nullptr;
            HRESULT shr = SHCreateStreamOnFileEx(
                path, STGM_READ | STGM_SHARE_DENY_NONE,
                0, FALSE, nullptr, &pStream);
            HostLog(L"  SHCreateStreamOnFileEx -> 0x%08lX", static_cast<long>(shr));
            if (SUCCEEDED(shr))
            {
                HRESULT ihr = pInit->Initialize(pStream, STGM_READ);
                HostLog(L"  IInitializeWithStream::Initialize -> 0x%08lX",
                        static_cast<long>(ihr));
                pStream->Release();
                initialized = SUCCEEDED(ihr);
                if (!initialized) hr = ihr;
            }
            else
            {
                hr = shr;
            }
            pInit->Release();
        }
    }

    // IInitializeWithItem is the modern Shell-item-based initializer.
    // Outlook's MSG preview handler implements only this one (it rejects
    // both IInitializeWithFile and IInitializeWithStream with E_NOINTERFACE).
    if (!initialized)
    {
        IInitializeWithItem* pInit = nullptr;
        HRESULT qhr = pUnk->QueryInterface(__uuidof(IInitializeWithItem),
                                           reinterpret_cast<void**>(&pInit));
        HostLog(L"  QI IInitializeWithItem -> 0x%08lX", static_cast<long>(qhr));
        if (SUCCEEDED(qhr))
        {
            IShellItem* pItem = nullptr;
            HRESULT shr = SHCreateItemFromParsingName(
                path, nullptr, __uuidof(IShellItem),
                reinterpret_cast<void**>(&pItem));
            HostLog(L"  SHCreateItemFromParsingName -> 0x%08lX",
                    static_cast<long>(shr));
            if (SUCCEEDED(shr))
            {
                HRESULT ihr = pInit->Initialize(pItem, STGM_READ);
                HostLog(L"  IInitializeWithItem::Initialize -> 0x%08lX",
                        static_cast<long>(ihr));
                pItem->Release();
                initialized = SUCCEEDED(ihr);
                if (!initialized) hr = ihr;
            }
            else
            {
                hr = shr;
            }
            pInit->Release();
        }
    }

    if (!initialized)
    {
        HostLog(L"  no Initialize* interface succeeded");
        // Drop the site we installed earlier before releasing the handler.
        {
            IObjectWithSite* pOws = nullptr;
            if (SUCCEEDED(pPH->QueryInterface(__uuidof(IObjectWithSite),
                                               reinterpret_cast<void**>(&pOws))))
            {
                pOws->SetSite(nullptr);
                pOws->Release();
            }
        }
        pPH->Release();
        pUnk->Release();
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    RECT rc = {};
    GetClientRect(g_state.hwndRender, &rc);
    HostLog(L"  render client rect = (%ld,%ld)-(%ld,%ld)",
            rc.left, rc.top, rc.right, rc.bottom);

    hr = pPH->SetWindow(g_state.hwndRender, &rc);
    HostLog(L"  IPreviewHandler::SetWindow -> 0x%08lX", static_cast<long>(hr));
    if (SUCCEEDED(hr))
    {
        hr = pPH->DoPreview();
        HostLog(L"  IPreviewHandler::DoPreview -> 0x%08lX", static_cast<long>(hr));
    }

    if (FAILED(hr))
    {
        // Drop the site we installed before releasing the handler.
        IObjectWithSite* pOws = nullptr;
        if (SUCCEEDED(pPH->QueryInterface(__uuidof(IObjectWithSite),
                                           reinterpret_cast<void**>(&pOws))))
        {
            pOws->SetSite(nullptr);
            pOws->Release();
        }
        pPH->Release();
        pUnk->Release();
        return hr;
    }

    g_state.pHandlerUnk = pUnk;
    g_state.pHandler    = pPH;

    // Force a clean repaint of the render pane and all children so the new
    // handler's window paints over any leftover area from the previous one
    // immediately (instead of waiting for the next WM_PAINT cycle).
    RedrawWindow(g_state.hwndRender, nullptr, nullptr,
                 RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);

    HostLog(L"  TryLoadHandlerSta SUCCESS");
    return S_OK;
}

// Wrapper used when the caller doesn't want to handle preview-handler
// failure itself: shows the fallback panel on any failure HRESULT and
// always returns S_OK to the pipe (TC keeps the Lister window either way).
static HRESULT LoadHandlerSta(LPCWSTR path)
{
    HRESULT hr = TryLoadHandlerSta(path);
    if (FAILED(hr))
        ShowFallbackSta(path, hr);
    return S_OK;
}

// ---------------------------------------------------------------------------
// Top-level LOAD dispatcher. Picks quick vs full mode by extension+config and
// hands off to the matching loader. On full-mode failure we degrade silently
// to quick mode so the user always gets *some* preview.
// ---------------------------------------------------------------------------

// Threshold above which we substitute the path with a junction-based short
// alias. MAX_PATH-1 = 259 is the limit that several Office consumers
// enforce internally; subtracting another 8 characters gives us margin
// for the "8.3" filename suffix some Win32 APIs append.
static constexpr size_t kLongPathAliasThreshold = MAX_PATH - 9;

// Try to remove a junction directory.  Best-effort: on failure (typical
// cause is Office having a directory-change-notification handle on the
// target) the path is added to the stale list for later retry instead of
// being silently leaked.
static void TryRemoveJunctionSta(const std::wstring& dir)
{
    if (dir.empty()) return;
    if (RemoveDirectoryW(dir.c_str()))
    {
        HostLog(L"  TryRemoveJunction OK: '%s'", dir.c_str());
        return;
    }
    DWORD err = GetLastError();
    HostLog(L"  TryRemoveJunction err=%lu — deferring '%s'", err, dir.c_str());
    g_state.staleJunctionDirs.push_back(dir);
}

// Walk staleJunctionDirs, dropping ones we can now remove.
static void RetryStaleJunctionsSta()
{
    auto& stale = g_state.staleJunctionDirs;
    for (auto it = stale.begin(); it != stale.end(); )
    {
        if (RemoveDirectoryW(it->c_str()))
        {
            HostLog(L"  Retry junction cleanup OK: '%s'", it->c_str());
            it = stale.erase(it);
        }
        else
        {
            ++it;
        }
    }
}

// Remove the currently-installed junction (if any) and retry any
// previously-deferred ones.  Caller must arrange that anything that
// held the target — Word document, Excel workbook, PowerPoint
// presentation, preview handler — has been closed before this runs.
static void CleanupActiveJunctionSta()
{
    if (!g_state.activeJunctionDir.empty())
    {
        std::wstring dir = std::move(g_state.activeJunctionDir);
        g_state.activeJunctionDir.clear();
        TryRemoveJunctionSta(dir);
    }
    RetryStaleJunctionsSta();
}

// If `origPath` is long enough to trip a consumer's MAX_PATH check,
// create a directory junction in %TEMP% pointing at its parent folder
// and return a short alias path inside that junction.  On any failure
// returns origPath unchanged (the caller will then get whatever error
// the underlying consumer raises).
//
// Side effects: on success, g_state.activeJunctionDir is set to the
// junction directory's path so CleanupActiveJunctionSta can remove it.
static std::wstring MakeLongPathAliasSta(LPCWSTR origPath)
{
    const size_t len = wcslen(origPath);
    if (len < kLongPathAliasThreshold) return origPath;

    LPCWSTR sep = wcsrchr(origPath, L'\\');
    if (!sep) return origPath;                              // no parent, give up
    const std::wstring parentDir(origPath, sep - origPath);
    const std::wstring filename(sep + 1);

    // Resolve %TEMP% dynamically so a long-path %TEMP% still works.
    DWORD needed = GetEnvironmentVariableW(L"TEMP", nullptr, 0);
    if (needed == 0) return origPath;
    std::wstring tempDir(needed, L'\0');
    DWORD got = GetEnvironmentVariableW(L"TEMP", tempDir.data(), needed);
    if (got == 0 || got >= needed) return origPath;
    tempDir.resize(got);
    if (!tempDir.empty() && tempDir.back() == L'\\') tempDir.pop_back();

    wchar_t suffix[64] = {};
    swprintf_s(suffix, ARRAYSIZE(suffix), L"\\TCOV_%lu_%llu",
               GetCurrentProcessId(),
               static_cast<unsigned long long>(GetTickCount64()));
    const std::wstring junctionDir = tempDir + suffix;

    if (!CreateJunctionSta(junctionDir, parentDir))
        return origPath;                                    // fallback — caller fails noisily

    g_state.activeJunctionDir = junctionDir;
    const std::wstring alias = junctionDir + L"\\" + filename;
    HostLog(L"  MakeLongPathAlias: %zu-char path -> %zu-char alias via '%s'",
            len, alias.size(), junctionDir.c_str());
    return alias;
}

// Sweep %TEMP% for orphaned TCOV_<pid>_<tick> junctions left behind by
// previous host instances that exited abnormally (TC killed during
// shutdown, host crash, etc.).  Junctions whose owner PID is still
// alive are skipped — they belong to a concurrently-running host.
// Called once at host start-up.
static void CleanupOrphanedJunctionsAtStartup()
{
    DWORD needed = GetEnvironmentVariableW(L"TEMP", nullptr, 0);
    if (needed == 0) return;
    std::wstring tempDir(needed, L'\0');
    DWORD got = GetEnvironmentVariableW(L"TEMP", tempDir.data(), needed);
    if (got == 0 || got >= needed) return;
    tempDir.resize(got);
    if (!tempDir.empty() && tempDir.back() == L'\\') tempDir.pop_back();

    const std::wstring pattern = tempDir + L"\\TCOV_*";
    WIN32_FIND_DATAW fd = {};
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE) return;

    int removed = 0, skipped = 0;
    do
    {
        // Only directories that are reparse points (our junctions).
        if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) continue;
        if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT)) continue;
        if (wcsncmp(fd.cFileName, L"TCOV_", 5) != 0) continue;

        // Parse owner PID from the name.
        DWORD pid = 0;
        if (swscanf_s(fd.cFileName + 5, L"%lu_", &pid) != 1 || pid == 0)
            continue;
        if (pid == GetCurrentProcessId())
            continue;                                   // ours (impossible at start-up, defensive)

        // If the owning host is still running, leave its junction alone.
        HANDLE hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
        if (hProc)
        {
            DWORD exitCode = 0;
            bool stillAlive = GetExitCodeProcess(hProc, &exitCode)
                              && exitCode == STILL_ACTIVE;
            CloseHandle(hProc);
            if (stillAlive) { ++skipped; continue; }
        }

        const std::wstring full = tempDir + L"\\" + fd.cFileName;
        if (RemoveDirectoryW(full.c_str()))
        {
            HostLog(L"  StartupSweep: removed orphan (pid %lu): '%s'",
                    pid, full.c_str());
            ++removed;
        }
        else
        {
            HostLog(L"  StartupSweep: RemoveDirectoryW(%s) failed err=%lu",
                    full.c_str(), GetLastError());
        }
    } while (FindNextFileW(hFind, &fd));

    FindClose(hFind);

    if (removed || skipped)
        HostLog(L"  StartupSweep: removed=%d, skipped (live pid)=%d",
                removed, skipped);
}

// Internal worker for LoadFileSta. Mode is supplied explicitly so the
// user-initiated WM_HOST_SWITCH_MODE handler can force the opposite of
// the INI-configured default for the current preview.
//
// `allowAutoFallback` controls the implicit quick→full retry for Office
// files whose quick-mode load failed.  The initial LOAD path (LoadFileSta,
// via WM_HOST_LOAD) passes `true` so SharePoint multi-tenant documents
// "just work".  WM_HOST_SWITCH_MODE passes `false` — the user has
// explicitly asked for a specific mode and an implicit retry would either
// (a) bounce them back to the mode they were trying to leave, or
// (b) silently ignore their click.  When the user-initiated switch lands
// on the fallback panel, the mode-switch button updates to point at the
// other mode so the user can flip back deliberately.
static HRESULT LoadFileWithModeSta(LPCWSTR origPath, AppKind app, Mode mode,
                                   bool allowAutoFallback = true)
{
    if (g_state.loadingInProgress)
    {
        HostLog(L"LoadFileWithModeSta: ignored — previous load still running");
        return E_ABORT;
    }
    g_state.loadingInProgress = true;

    // currentFile stores the user's ORIGINAL path; the mode-switch
    // handler uses it to re-LOAD and recreates the alias from scratch.
    g_state.currentFile       = origPath;
    g_state.currentFileApp    = app;
    g_state.currentLoadedMode = mode;

    // Move the previous junction aside, but do NOT remove it yet — the
    // previous consumer (Word document, Excel workbook, ...) is still
    // alive and may hold a directory-change-notification handle on the
    // junction, which would make RemoveDirectoryW fail.  We retry the
    // removal at the end of this function, after the loader below has
    // closed the previous document.
    std::wstring prevJunctionDir = std::move(g_state.activeJunctionDir);
    g_state.activeJunctionDir.clear();                  // defensive (move already cleared)

    // For paths above the MAX_PATH ceiling that Office consumers enforce
    // internally, substitute a short alias path inside a freshly-created
    // %TEMP%\TCOV_… junction.  Everything downstream sees the alias.
    const std::wstring aliasStorage = MakeLongPathAliasSta(origPath);
    LPCWSTR path = aliasStorage.c_str();

    HostLog(L"LoadFileWithModeSta: app=%d mode=%s%s",
            static_cast<int>(app),
            mode == Mode::Full ? L"full" : L"quick",
            g_state.activeJunctionDir.empty() ? L"" : L" (via junction)");

    HRESULT result          = E_FAIL;
    bool    fellThroughFull = false;        // full mode failed → try quick

    // If the file has MOTW and we're about to open it in full mode, skip
    // directly to fallback. Opening a MOTW-blocked file in the real Office
    // app would load it in an editable state (not Protected View) because
    // ReadOnly=True bypasses Protected View. Show the fallback instead.
    {
        const MotwInfo motw = ReadFileZoneInfo(path);
        if (motw.zoneId >= 3 && mode == Mode::Full)
        {
            HostLog(L"  MOTW detected (ZoneId=%d), skipping full mode — showing fallback",
                    motw.zoneId);
            // Tear down any previously-embedded Office app so the fallback
            // panel isn't visually hidden behind it.  This is the same
            // light-weight cleanup that the quick-mode fall-through path
            // below performs: documents close, but the Office processes
            // stay alive so the next non-MOTW Full-mode load is fast.
            if (g_state.hwndWordApp)  { CloseWordDocumentSta();    DetachWordWindowSta(); }
            if (g_state.hwndExcelApp) { CloseExcelWorkbookSta();   DetachExcelWindowSta(); }
            if (g_state.hwndPptApp)   { ClosePptPresentationSta(); DetachPptWindowSta(); }
            // Show the original user-visible path, not the junction alias.
            ShowFallbackSta(origPath, E_FAIL);
            // Clean up the previous LOAD's junction (we didn't reach the
            // function's normal cleanup at the bottom because of the
            // early return); also retry any older stale junctions.
            TryRemoveJunctionSta(prevJunctionDir);
            RetryStaleJunctionsSta();
            g_state.loadingInProgress = false;
            return S_OK;
        }
    }

    if (mode == Mode::Full)
    {
        HRESULT hr = E_FAIL;
        LPCWSTR appName = L"?";
        if (app == AppKind::Word)
        {
            appName = L"Word";
            hr = LoadWordFullSta(path);
        }
        else if (app == AppKind::Excel)
        {
            appName = L"Excel";
            hr = LoadExcelFullSta(path);
        }
        else if (app == AppKind::PowerPoint)
        {
            appName = L"PowerPoint";
            hr = LoadPowerPointFullSta(path);
        }

        if (app == AppKind::Word || app == AppKind::Excel || app == AppKind::PowerPoint)
        {
            if (SUCCEEDED(hr))
            {
                UpdateModeButtonSta();
                // Office apps may asynchronously reposition their embedded
                // window after load (e.g. in response to WM_PARENTNOTIFY),
                // covering the button.  Schedule a deferred re-raise.
                if (g_state.hwndRender)
                    SetTimer(g_state.hwndRender, kModeButtonZTimerId, 400, nullptr);
                result = S_OK;
            }
            else
            {
                HostLog(L"  %s full-mode failed (0x%08lX) — falling back to quick",
                        appName, static_cast<long>(hr));
                // Hard-reset all three app slots so quick mode starts from a
                // clean state. Each Unload is a no-op if its slot was empty.
                UnloadWordFullSta(true);
                UnloadExcelFullSta(true);
                UnloadPptFullSta(true);
                // Failed full mode → actual loaded mode will be quick.
                g_state.currentLoadedMode = Mode::Quick;
                fellThroughFull = true;
            }
        }
    }

    if (mode != Mode::Full || fellThroughFull)
    {
        // Quick mode (or full-mode fallback path). If any Office app's window
        // is still parented under the render pane, detach it first.
        // Detach alone is not enough — we must also close the document so the
        // Office app does not hold a lock on the file. Otherwise a rapid
        // quick→full→quick switch can leave the document open in the background
        // and the next full-mode open may fail or behave unpredictably.
        if (g_state.hwndWordApp)  { CloseWordDocumentSta(); DetachWordWindowSta(); }
        if (g_state.hwndExcelApp) { CloseExcelWorkbookSta(); DetachExcelWindowSta(); }
        if (g_state.hwndPptApp)   { ClosePptPresentationSta(); DetachPptWindowSta(); }

        HRESULT quickHr = TryLoadHandlerSta(path);
        result = S_OK;

        // Auto-fallback quick → full for Office files when the preview
        // handler refused them.  Typical trigger is SharePoint documents
        // synced from non-primary Microsoft 365 tenants: the prevhost.exe
        // surrogate the Office preview handler runs in can't authenticate
        // cross-tenant, so quick mode fails with E_FAIL / E_NOTIMPL.  The
        // real Office app, running as the user with full auth tokens, can
        // open the same file fine.  (The same limit affects Explorer's
        // Alt+P pane — it's an Office surrogate restriction, not a
        // TCOfficeView issue — so the auto-fallback is the user-visible
        // fix.)
        //
        // Gated on six conditions to keep the implicit mode switch
        // predictable:
        //   1) The caller permits it (allowAutoFallback).  WM_HOST_LOAD
        //      passes true (initial open should "just work" for SharePoint
        //      cross-tenant files); WM_HOST_SWITCH_MODE passes false (the
        //      user just clicked → Quick and an implicit bounce back to
        //      Full would feel like the button did nothing).
        //   2) The file is Office (Word / Excel / PowerPoint).
        //   3) The user has configured `quick-switchable` for that app —
        //      i.e. they have already opted into mode switching being a
        //      sensible thing to happen for this app type.  `quick`
        //      (non-switchable) means the user explicitly wants quick
        //      only, even on failure.
        //   4) [AutoFallback] for that app is true in the INI (default).
        //   5) We didn't already arrive here from a failed full-mode
        //      attempt (would loop forever).
        //   6) The file is NOT Mark-of-the-Web blocked.  MOTW files have
        //      their own dedicated fallback panel with the "Unblock"
        //      button — auto-fallback to full would bypass that UX (and
        //      worse, open the file editably because ReadOnly=True at the
        //      Office app level skips Protected View).  Same MOTW guard
        //      that the explicit-full path enforces earlier in this
        //      function.
        const bool isOfficeApp =
            (app == AppKind::Word || app == AppKind::Excel ||
             app == AppKind::PowerPoint);
        const bool isMotwBlocked =
            isOfficeApp && (ReadFileZoneInfo(origPath).zoneId >= 3);
        const bool fallbackPermitted =
            allowAutoFallback &&
            isOfficeApp &&
            !fellThroughFull &&
            !isMotwBlocked &&
            SelectMode(app) == Mode::QuickSwitchable &&
            IsAutoFallbackEnabled(app);

        if (FAILED(quickHr) && fallbackPermitted)
        {
            HostLog(L"  quick mode failed (0x%08lX) for Office file — "
                    L"auto-fallback to full",
                    static_cast<long>(quickHr));
            HRESULT fullHr = E_FAIL;
            LPCWSTR appName = L"?";
            if (app == AppKind::Word)
            {
                appName = L"Word";
                fullHr = LoadWordFullSta(path);
            }
            else if (app == AppKind::Excel)
            {
                appName = L"Excel";
                fullHr = LoadExcelFullSta(path);
            }
            else if (app == AppKind::PowerPoint)
            {
                appName = L"PowerPoint";
                fullHr = LoadPowerPointFullSta(path);
            }

            if (SUCCEEDED(fullHr))
            {
                HostLog(L"  %s full-mode auto-fallback SUCCESS", appName);
                g_state.currentLoadedMode = Mode::Full;
                if (g_state.hwndRender)
                    SetTimer(g_state.hwndRender, kModeButtonZTimerId,
                             400, nullptr);
            }
            else
            {
                HostLog(L"  %s full-mode auto-fallback also failed (0x%08lX) "
                        L"— showing fallback panel with quick-mode error",
                        appName, static_cast<long>(fullHr));
                UnloadWordFullSta(true);
                UnloadExcelFullSta(true);
                UnloadPptFullSta(true);
                // Show the quick-mode HRESULT, not the full-mode one — it
                // describes the actual rejection of the file's content
                // (MOTW, IRM, missing handler, …) rather than a generic
                // Office startup failure.
                ShowFallbackSta(path, quickHr);
            }
        }
        else if (FAILED(quickHr))
        {
            // Non-Office file, fallback disabled, or full already failed
            // for this load — just show the fallback panel.
            ShowFallbackSta(path, quickHr);
        }

        UpdateModeButtonSta();
        // Quick-mode preview handlers may also do async window layout; a
        // shorter delay is enough since they don't need Office startup time.
        if (g_state.hwndRender)
            SetTimer(g_state.hwndRender, kModeButtonZTimerId, 200, nullptr);
    }

    // The loader above has now closed any previous consumer's document /
    // released its file handles.  Attempt to remove the old junction
    // (which we stashed at function entry); if Office STILL has a handle
    // on it, it joins the stale list for retry on a later cleanup
    // opportunity (next LOAD, WM_HOST_CLOSE, or the next host process's
    // startup sweep).
    TryRemoveJunctionSta(prevJunctionDir);
    RetryStaleJunctionsSta();

    g_state.loadingInProgress = false;
    return result;
}

// Public LoadFile entry point — uses the INI-configured mode. This is the
// path taken by the WM_HOST_LOAD pipe command (one per file shown in the
// Lister), so switching files always resets to the configured default and
// the user's button click does not carry over.
static HRESULT LoadFileSta(LPCWSTR path)
{
    AppKind app = ClassifyByExtension(path);
    Mode    cfg = SelectMode(app);              // may include the Switchable suffix
    return LoadFileWithModeSta(path, app, BaseMode(cfg));
}

static void ResizeHandlerSta(int w, int h)
{
    if (g_state.hwndRender)
    {
        SetWindowPos(g_state.hwndRender, nullptr, 0, 0, w, h,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (g_state.pHandler)
    {
        RECT rc = { 0, 0, w, h };
        g_state.pHandler->SetRect(&rc);
    }
    if (g_state.hwndFallback)
    {
        SetWindowPos(g_state.hwndFallback, nullptr, 0, 0, w, h,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }
    // The unblock button is positioned relative to the bottom edge; without
    // this it would stay at its original coordinates and end up off-screen
    // (or covered) when the user shrinks the Lister window.
    LayoutUnblockButtonSta(w, h);
    ResizeOfficeFullSta(w, h);

    // Keep the close-guard stretched across the full width of the render pane.
    // Must happen BEFORE UpdateModeButtonSta so that the subsequent HWND_TOP
    // raise of the mode button lands above the guard (not below it).
    if (g_state.hwndCloseGuard)
    {
        UINT dpi = GetDpiForWindow(g_state.hwndRender);
        int gh = ScaleForDpi(40, dpi);
        SetWindowPos(g_state.hwndCloseGuard, HWND_TOP,
                     0, 0, w, gh,
                     SWP_NOACTIVATE | SWP_SHOWWINDOW);
    }
    UpdateModeButtonSta();                       // reposition overlay button above close guard
}

// ---------------------------------------------------------------------------
// STA window procedure
// ---------------------------------------------------------------------------

static BOOL PipeWriteUtf16(HANDLE h, LPCWSTR text);    // defined below

static LRESULT CALLBACK StaWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_HOST_LOAD:
        {
            wchar_t* path = reinterpret_cast<wchar_t*>(lp);
            HRESULT hr = LoadFileSta(path);
            if (SUCCEEDED(hr))
            {
                PipeWriteUtf16(g_state.hPipe, L"OK\n");
            }
            else
            {
                wchar_t errMsg[96] = {};
                _snwprintf_s(errMsg, ARRAYSIZE(errMsg), _TRUNCATE,
                             L"ERR LoadFile hr=0x%08lX\n",
                             static_cast<long>(hr));
                PipeWriteUtf16(g_state.hPipe, errMsg);
            }
            delete[] path;
            return 0;
        }
        case WM_HOST_RESIZE:
        {
            int w = static_cast<int>(LOWORD(wp));
            int h = static_cast<int>(HIWORD(wp));
            ResizeHandlerSta(w, h);
            return 0;
        }
        case WM_HOST_SWITCH_MODE:
        {
            // User clicked the mode-switch overlay. Re-load the current
            // file in the opposite mode (this is per-preview only — the
            // next LOAD command from the plugin DLL goes back to the
            // INI-configured default via LoadFileSta).
            if (g_state.currentFile.empty())
            {
                HostLog(L"WM_HOST_SWITCH_MODE: ignored (no file currently loaded)");
                return 0;
            }
            Mode target = (g_state.currentLoadedMode == Mode::Quick)
                            ? Mode::Full
                            : Mode::Quick;
            HostLog(L"WM_HOST_SWITCH_MODE: switching to %s",
                    target == Mode::Full ? L"full" : L"quick");
            // currentFile is owned by g_state; copy locally because
            // LoadFileWithModeSta overwrites g_state.currentFile early on
            // (which would invalidate the c_str() pointer mid-call).
            std::wstring path = g_state.currentFile;
            // allowAutoFallback=false: the user explicitly asked for `target`,
            // so the implicit quick→full retry path must NOT kick in here.
            // Otherwise a user who clicks → Quick on a file that auto-fell-
            // back to full at load time would simply bounce back to full
            // and the click would look like a no-op.  When `target` is Quick
            // and the preview handler still refuses the file, ShowFallbackSta
            // shows the fallback panel and the mode-switch button auto-flips
            // back to → Full so the user can return deliberately.
            HRESULT hr = LoadFileWithModeSta(path.c_str(),
                                             g_state.currentFileApp,
                                             target,
                                             /*allowAutoFallback=*/false);
            if (FAILED(hr))
            {
                HostLog(L"WM_HOST_SWITCH_MODE: LoadFileWithModeSta failed hr=0x%08lX",
                        static_cast<long>(hr));
            }
            return 0;
        }
        case WM_HOST_UNBLOCK_AND_RELOAD:
        {
            // User clicked the MOTW Unblock button in the fallback panel.
            // Strip the file's Zone.Identifier alternate data stream and
            // re-attempt the LOAD with the same mode that was active
            // when the fallback was shown.
            if (g_state.currentFile.empty())
            {
                HostLog(L"WM_HOST_UNBLOCK_AND_RELOAD: ignored (no current file)");
                return 0;
            }
            std::wstring path = g_state.currentFile;     // local copy; reload overwrites
            const AppKind app  = g_state.currentFileApp;
            const Mode    mode = g_state.currentLoadedMode;

            bool ok = UnblockFileSta(path.c_str());
            HostLog(L"WM_HOST_UNBLOCK_AND_RELOAD: DeleteFile('%s:Zone.Identifier') -> %s",
                    path.c_str(), ok ? L"OK" : L"FAILED");

            if (!ok)
            {
                // ADS removal failed (file read-only, antivirus locked it, …).
                // Leave the fallback panel as-is so the user can see what's
                // going on; no further action without their input.
                return 0;
            }

            HRESULT hr = LoadFileWithModeSta(path.c_str(), app, mode);
            if (FAILED(hr))
            {
                HostLog(L"WM_HOST_UNBLOCK_AND_RELOAD: reload failed hr=0x%08lX",
                        static_cast<long>(hr));
            }
            return 0;
        }
        case WM_HOST_CLOSE:
        {
            HostLog(L"WM_HOST_CLOSE");
            UnloadHandlerSta();
            // Quit any Office app that's still alive. Each is a no-op if
            // its slot is empty.
            UnloadWordFullSta(true);
            UnloadExcelFullSta(true);
            UnloadPptFullSta(true);
            // Remove any long-path junction we created in %TEMP%.
            // Office docs have already been closed above, so the
            // target's file handles are released.  Anything that's
            // still stale at this point will be picked up by the
            // next host launch's CleanupOrphanedJunctionsAtStartup.
            CleanupActiveJunctionSta();
            if (!g_state.staleJunctionDirs.empty())
                HostLog(L"  WM_HOST_CLOSE: %zu junction(s) still stale, "
                        L"deferring to next host's startup sweep",
                        g_state.staleJunctionDirs.size());
            DestroyRenderWindowSta();
            PipeWriteUtf16(g_state.hPipe, L"OK\n");
            PostQuitMessage(0);
            return 0;
        }
    }
    return DefWindowProcW(hWnd, msg, wp, lp);
}

// ---------------------------------------------------------------------------
// Overlapped pipe helpers (sync-wrapped: appear blocking to the caller).
// The pipe is opened with FILE_FLAG_OVERLAPPED so the plugin can use a
// timeout on ConnectNamedPipe; that requires every read/write to supply
// its own OVERLAPPED structure.
// ---------------------------------------------------------------------------

static BOOL PipeReadSync(HANDLE h, void* buf, DWORD size, DWORD* outBytes)
{
    OVERLAPPED ov = {};
    ov.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    BOOL ok = ReadFile(h, buf, size, nullptr, &ov);
    if (!ok && GetLastError() != ERROR_IO_PENDING && GetLastError() != ERROR_MORE_DATA)
    {
        CloseHandle(ov.hEvent);
        return FALSE;
    }
    ok = GetOverlappedResult(h, &ov, outBytes, TRUE);
    CloseHandle(ov.hEvent);
    return ok;
}

static BOOL PipeWriteSync(HANDLE h, const void* data, DWORD size)
{
    OVERLAPPED ov = {};
    ov.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    BOOL ok = WriteFile(h, data, size, nullptr, &ov);
    if (!ok && GetLastError() != ERROR_IO_PENDING)
    {
        CloseHandle(ov.hEvent);
        return FALSE;
    }
    DWORD written = 0;
    ok = GetOverlappedResult(h, &ov, &written, TRUE);
    CloseHandle(ov.hEvent);
    return ok && written == size;
}

static BOOL PipeWriteUtf16(HANDLE h, LPCWSTR text)
{
    DWORD bytes = static_cast<DWORD>(wcslen(text) * sizeof(wchar_t));
    return PipeWriteSync(h, text, bytes);
}

// ---------------------------------------------------------------------------
// Pipe reader worker
// ---------------------------------------------------------------------------

static DWORD WINAPI PipeReaderThread(LPVOID)
{
    constexpr DWORD kBufWChars = 32 * 1024;
    auto* buf = new wchar_t[kBufWChars];
    for (;;)
    {
        DWORD bytesRead = 0;
        if (!PipeReadSync(g_state.hPipe, buf,
                          (kBufWChars - 1) * sizeof(wchar_t), &bytesRead))
            break;
        if (bytesRead == 0) break;

        DWORD wchars = bytesRead / sizeof(wchar_t);
        buf[wchars] = L'\0';

        // Strip trailing CR / LF that the plugin appends.
        while (wchars > 0 && (buf[wchars - 1] == L'\n' || buf[wchars - 1] == L'\r'))
            buf[--wchars] = L'\0';

        if (wcsncmp(buf, L"LOAD ", 5) == 0)
        {
            // Copy the path onto the heap; ownership transfers to the WndProc
            // handler on the STA thread, which deletes it after processing.
            size_t pathChars = wcslen(buf + 5);
            auto* heapPath = new wchar_t[pathChars + 1];
            wcscpy_s(heapPath, pathChars + 1, buf + 5);
            if (!PostMessageW(g_state.hwndSta, WM_HOST_LOAD,
                              0, reinterpret_cast<LPARAM>(heapPath)))
            {
                delete[] heapPath;
                PipeWriteUtf16(g_state.hPipe, L"ERR PostMessage failed\n");
            }
            // STA writes the OK / ERR response itself once the load completes.
        }
        else if (wcsncmp(buf, L"RESIZE ", 7) == 0)
        {
            int w = 0, h = 0;
            if (swscanf_s(buf + 7, L"%d %d", &w, &h) == 2 && w > 0 && h > 0)
            {
                WPARAM wp = MAKELONG(w, h);
                PostMessageW(g_state.hwndSta, WM_HOST_RESIZE, wp, 0);
            }
            // RESIZE is fire-and-forget — no response.
        }
        else if (wcsncmp(buf, L"CLOSE", 5) == 0)
        {
            break;
        }
        else
        {
            PipeWriteUtf16(g_state.hPipe, L"ERR unknown command\n");
        }
    }

    delete[] buf;

    // Kick the STA out of its message loop (graceful or unexpected pipe death).
    PostMessageW(g_state.hwndSta, WM_HOST_CLOSE, 0, 0);
    return 0;
}

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------

static bool ParseArgs(int argc, wchar_t** argv, HWND* outHwnd, std::wstring* outPipe)
{
    *outHwnd = nullptr;
    outPipe->clear();
    for (int i = 1; i + 1 < argc; ++i)
    {
        if (wcscmp(argv[i], L"--hwnd") == 0)
            *outHwnd = reinterpret_cast<HWND>(static_cast<INT_PTR>(_wtoi64(argv[i + 1])));
        else if (wcscmp(argv[i], L"--pipe") == 0)
            *outPipe = argv[i + 1];
    }
    return *outHwnd != nullptr && !outPipe->empty();
}

static HWND CreateStaWindow(HINSTANCE hInst)
{
    static const wchar_t* kClass = L"TCOfficeViewHostSta";
    WNDCLASSW wc = {};
    wc.lpfnWndProc   = StaWndProc;
    wc.hInstance     = hInst;
    wc.lpszClassName = kClass;
    RegisterClassW(&wc);
    return CreateWindowExW(0, kClass, L"", 0, 0, 0, 0, 0,
                           HWND_MESSAGE, nullptr, hInst, nullptr);
}

int wmain(int argc, wchar_t** argv)
{
    // Initialize the logging critical section before any HostLog() call.
    // This eliminates the TOCTOU race that existed when the CS was lazily
    // initialised inside HostLog on the first call.
    InitializeCriticalSection(&g_logCs);

    // Must run before the first HostLog() call so g_logPath is populated.
    LoadConfig();

    std::wstring pipeName;
    if (!ParseArgs(argc, argv, &g_state.hwndPluginChild, &pipeName))
    {
        HostLog(L"wmain: bad args (need --hwnd <h> --pipe <name>)");
        return 1;
    }
    HostLog(L"wmain: plugin-child=0x%p pipe='%s'",
            g_state.hwndPluginChild, pipeName.c_str());

    // Sweep %TEMP% for junctions left behind by host instances that
    // exited abnormally.  Cheap (one FindFirstFile + a few PID checks)
    // and keeps long-lived TC sessions from accumulating orphans.
    CleanupOrphanedJunctionsAtStartup();

    // Optional: regenerate the available-preview-handlers discovery
    // report on a low-priority background thread.  Skipped entirely
    // when the user hasn't enabled it ([PreviewHandlers] ReportPath
    // is empty) so the host doesn't pay the registry-walk cost on
    // every Lister session.
    if (!g_handlersReportPath.empty())
    {
        if (HANDLE hReportThread = CreateThread(nullptr, 0,
                                                AvailableHandlersReportThread,
                                                nullptr, 0, nullptr))
        {
            CloseHandle(hReportThread);
        }
    }

    HRESULT hr = CoInitializeEx(nullptr,
                                COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    HostLog(L"  CoInitializeEx -> 0x%08lX", static_cast<long>(hr));
    if (FAILED(hr))
        return 2;

    // Establish a process-wide COM security blanket BEFORE the first COM
    // activation.  This matches Microsoft's PreviewHandler hosting sample
    // and is mandatory for hosting handlers that run in the low-integrity
    // prevhost.exe surrogate (notably Outlook's MSG handler): without it,
    // cross-process calls between us and the surrogate fall back to a
    // restrictive default that prevents the handler from completing
    // DoPreview (it returns E_FAIL).  Once any interface is marshaled the
    // default security is locked in, so this call must precede every
    // CoCreateInstance — putting it here in wmain guarantees that.
    //
    // RPC_C_IMP_LEVEL_IMPERSONATE lets the surrogate impersonate our
    // identity when it needs to open files on our behalf; that is the
    // setting Windows Explorer's preview pane uses.
    HRESULT hrSec = CoInitializeSecurity(
        nullptr,                           // pSecDesc (use default)
        -1,                                // cAuthSvc (negotiate)
        nullptr,                           // asAuthSvc
        nullptr,                           // pReserved1
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        nullptr,                           // pAuthList
        EOAC_NONE,
        nullptr);                          // pReserved3
    HostLog(L"  CoInitializeSecurity -> 0x%08lX", static_cast<long>(hrSec));

    // Job Object that owns any Office processes we spawn (see the
    // AssignOfficeProcessToJobSta calls in the per-app loaders). Set up
    // KILL_ON_JOB_CLOSE so that when this host process dies — through
    // any mechanism, including being killed during Windows shutdown —
    // the kernel terminates the Office processes too. Without this
    // safety net Excel in particular tends to get left as an orphan
    // and the system reports a "hung application" on restart.
    g_state.hOfficeJob = CreateJobObjectW(nullptr, nullptr);
    if (g_state.hOfficeJob)
    {
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION info = {};
        info.BasicLimitInformation.LimitFlags =
            JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
        if (!SetInformationJobObject(g_state.hOfficeJob,
                                     JobObjectExtendedLimitInformation,
                                     &info, sizeof(info)))
        {
            HostLog(L"  SetInformationJobObject failed err=%lu — closing job",
                    GetLastError());
            CloseHandle(g_state.hOfficeJob);
            g_state.hOfficeJob = nullptr;
        }
    }
    else
    {
        HostLog(L"  CreateJobObject failed err=%lu — Office processes will not be auto-killed",
                GetLastError());
    }

    // Pipe was created by the plugin with FILE_FLAG_OVERLAPPED. Open the
    // client end with the matching flag.
    g_state.hPipe = CreateFileW(pipeName.c_str(),
                                GENERIC_READ | GENERIC_WRITE,
                                0, nullptr, OPEN_EXISTING,
                                FILE_FLAG_OVERLAPPED, nullptr);
    HostLog(L"  CreateFile on pipe -> %s (err=%lu)",
            g_state.hPipe == INVALID_HANDLE_VALUE ? L"FAIL" : L"OK",
            GetLastError());
    if (g_state.hPipe == INVALID_HANDLE_VALUE)
    {
        CoUninitialize();
        return 3;
    }
    DWORD mode = PIPE_READMODE_MESSAGE;
    SetNamedPipeHandleState(g_state.hPipe, &mode, nullptr, nullptr);

    HINSTANCE hInst = GetModuleHandleW(nullptr);

    if (!CreateRenderWindowSta(hInst))
    {
        HostLog(L"  CreateRenderWindowSta failed — bailing");
        CloseHandle(g_state.hPipe);
        CoUninitialize();
        return 4;
    }

    g_state.hwndSta = CreateStaWindow(hInst);
    if (!g_state.hwndSta)
    {
        DestroyRenderWindowSta();
        CloseHandle(g_state.hPipe);
        CoUninitialize();
        return 5;
    }

    HANDLE hThread = CreateThread(nullptr, 0, PipeReaderThread, nullptr, 0, nullptr);
    if (!hThread)
    {
        DestroyWindow(g_state.hwndSta);
        DestroyRenderWindowSta();
        CloseHandle(g_state.hPipe);
        CoUninitialize();
        return 6;
    }

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    // Make sure the reader exits before we close the pipe handle it shares.
    HANDLE hPipe = g_state.hPipe;
    g_state.hPipe = INVALID_HANDLE_VALUE;
    CancelIoEx(hPipe, nullptr);
    WaitForSingleObject(hThread, 2000);
    CloseHandle(hThread);
    CloseHandle(hPipe);

    DestroyWindow(g_state.hwndSta);
    UnregisterClassW(L"TCOfficeViewHostSta", hInst);
    g_state.hwndSta = nullptr;

    // Clean up the Job Object handle so the kernel can release it.
    if (g_state.hOfficeJob)
    {
        CloseHandle(g_state.hOfficeJob);
        g_state.hOfficeJob = nullptr;
    }

    CoUninitialize();

    // Close the log file handle and tear down the logging critical section.
    // These are done after CoUninitialize because HostLog may still be called
    // during COM cleanup (e.g. Release on Office objects).
    if (g_hLog && g_hLog != INVALID_HANDLE_VALUE)
    {
        CloseHandle(g_hLog);
        g_hLog = nullptr;
    }
    DeleteCriticalSection(&g_logCs);
    return 0;
}
