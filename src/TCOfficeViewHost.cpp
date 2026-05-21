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
#include <shobjidl.h>
#include <shlwapi.h>
#include <shlobj.h>           // SHCreateDirectoryExW
#include <propsys.h>          // IInitializeWithFile, IInitializeWithStream
#include <objbase.h>
#include <oleauto.h>          // IDispatch helpers: SysAllocString, VARIANT, ...
#include <stdio.h>
#include <string>

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

// Per-application render mode. "Quick" hosts the Windows Preview Handler
// registered by Office (fast, low memory, web-layout rendering only).
// "Full" drives a hidden Word/Excel/PowerPoint instance via OLE Automation
// and reparents its main window into our render pane (slower, higher
// memory, but gives the full editing-mode rendering — for Word that means
// the page layout with headers, footers and page breaks).
enum class Mode { Quick, Full };

static std::wstring g_logPath;       // empty → diagnostic logging is disabled
static std::wstring g_fontFamily;    // empty → auto-pick from a fallback list
static int          g_fontSize = 12;
static Mode         g_modeWord       = Mode::Quick;
static Mode         g_modeExcel      = Mode::Quick;
static Mode         g_modePowerPoint = Mode::Quick;

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
    wchar_t path[MAX_PATH] = {};
    GetModuleFileNameW(nullptr, path, MAX_PATH);
    std::wstring s(path);
    auto pos = s.find_last_of(L"\\/");
    return (pos != std::wstring::npos) ? s.substr(0, pos) : L".";
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
        if (_wcsicmp(buf, L"full")  == 0) return Mode::Full;
        if (_wcsicmp(buf, L"quick") == 0) return Mode::Quick;
        return dflt;                          // empty or invalid → default
    };
    g_modeWord       = readMode(L"Word",       Mode::Quick);
    g_modeExcel      = readMode(L"Excel",      Mode::Quick);
    g_modePowerPoint = readMode(L"PowerPoint", Mode::Quick);
    return true;
}

static void LoadConfig()
{
    // 1) Per-user override under TC's standard config directory.
    wchar_t appdata[MAX_PATH] = {};
    if (GetEnvironmentVariableW(L"APPDATA", appdata, MAX_PATH))
    {
        std::wstring userIni =
            std::wstring(appdata) + L"\\GHISLER\\TCOfficeView.ini";
        if (LoadConfigFrom(userIni)) return;
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

    IDispatch*          pPptApp         = nullptr;
    IDispatch*          pPptPres        = nullptr;
    HWND                hwndPptApp      = nullptr;        // PowerPoint main window reparented into hwndRender
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

// Control ID for the mode-switch BUTTON child window of hwndRender —
// referenced from RenderWndProc's WM_COMMAND handler.
static constexpr UINT_PTR kModeButtonId       = 1001;
// Timer ID used to re-raise the mode button to HWND_TOP after a resize.
// Preview handlers and Office apps may asynchronously reposition their own
// windows in response to a resize (via a posted internal message), burying
// the button.  A short one-shot timer fires after the handler settles.
static constexpr UINT_PTR kModeButtonZTimerId = 1002;

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

static HRESULT FindPreviewHandlerClsid(LPCWSTR path, CLSID* out)
{
    LPCWSTR dot = wcsrchr(path, L'.');
    if (!dot || !dot[1]) return E_INVALIDARG;

    std::wstring ext = dot;                              // e.g. ".docx"
    std::wstring catGuid = GuidToBraces(kPreviewHandlerCategory);

    std::wstring sub = ext + L"\\shellex\\" + catGuid;
    std::wstring clsidStr;
    bool foundDirect = ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr);
    HostLog(L"FindPreviewHandlerClsid: ext='%s' direct=%s value='%s'",
            ext.c_str(),
            foundDirect ? L"yes" : L"no",
            clsidStr.c_str());

    if (!foundDirect)
    {
        std::wstring progId;
        if (!ReadDefaultString(HKEY_CLASSES_ROOT, ext.c_str(), progId) || progId.empty())
        {
            HostLog(L"FindPreviewHandlerClsid: no ProgID for '%s'", ext.c_str());
            return REGDB_E_CLASSNOTREG;
        }
        sub = progId + L"\\shellex\\" + catGuid;
        if (!ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
        {
            HostLog(L"FindPreviewHandlerClsid: progId='%s' has no preview handler",
                    progId.c_str());
            return REGDB_E_CLASSNOTREG;
        }
        HostLog(L"FindPreviewHandlerClsid: progId='%s' value='%s'",
                progId.c_str(), clsidStr.c_str());
    }
    return CLSIDFromString(clsidStr.c_str(), out);
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
            return 1;                          // preview handler paints

        case WM_COMMAND:
            // Click on the mode-switch overlay button. Bounce it to the STA
            // window so the actual mode flip runs on the COM apartment
            // (LoadXxxFullSta calls IDispatch::Invoke on Word/Excel/PowerPoint).
            if (HIWORD(wp) == BN_CLICKED && LOWORD(wp) == kModeButtonId)
            {
                PostMessageW(g_state.hwndSta, WM_HOST_SWITCH_MODE, 0, 0);
                return 0;
            }
            break;

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
                KillTimer(hWnd, kModeButtonZTimerId);
                UpdateModeButtonSta();
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

static bool CreateModeButtonSta(HINSTANCE hInst)
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
        hInst, nullptr);
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

// Refreshes the mode-switch button. Called after every LOAD and every
// resize. Hides the button when the current file type doesn't support
// the alternative mode (i.e. when neither Word, Excel nor PowerPoint
// recognise the extension — no quick→full switch makes sense for those).
static void UpdateModeButtonSta()
{
    if (!g_state.hwndModeButton) return;

    const AppKind app = g_state.currentFileApp;
    const bool showable = (app == AppKind::Word
                        || app == AppKind::Excel
                        || app == AppKind::PowerPoint);
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
    CreateModeButtonSta(hInst);

    return true;
}

// Forward declaration — defined later in the close-guard section.
static void DestroyCloseGuard();

static void DestroyRenderWindowSta()
{
    DestroyCloseGuard();   // must go before hwndRender is destroyed
    if (g_state.hwndRender)
    {
        // Cancel any pending deferred Z-raise timer before the window dies.
        KillTimer(g_state.hwndRender, kModeButtonZTimerId);
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

static std::wstring BuildFallbackText(LPCWSTR path, HRESULT hr)
{
    std::wstring text;

    if (hr == REGDB_E_CLASSNOTREG)
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
                     L"(HRESULT 0x%08lX). Office may be installed but the\r\n"
                     L"handler is broken or the file itself may be corrupted.\r\n",
                     static_cast<long>(hr));
        text += lead;
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

    WIN32_FILE_ATTRIBUTE_DATA fad = {};
    if (GetFileAttributesExW(path, GetFileExInfoStandard, &fad))
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
}

static bool ShowFallbackSta(LPCWSTR path, HRESULT hr)
{
    HideFallbackSta();
    if (!g_state.hwndRender) return false;

    RECT rc = {};
    GetClientRect(g_state.hwndRender, &rc);

    HWND hEdit = CreateWindowExW(
        0, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_VSCROLL |
        ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL | ES_LEFT,
        0, 0, rc.right - rc.left, rc.bottom - rc.top,
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

// Excel preview tweaks: set a fixed 100% zoom. We deliberately don't
// touch Application.WindowState — for a WS_CHILD Excel frame, both
// xlNormal and xlMaximized produce visual glitches (the latter blanks
// the ribbon entirely). Layout for the initial size is handled in
// LoadExcelFullSta by setting Application.Width/Height *before*
// Workbooks.Open; ResizeOfficeFullSta keeps them in sync on resize.
// Excel preview tweak: fixed 100% zoom. Per-window setting that does not
// modify the workbook on disk (we opened ReadOnly).
//
// Known limit: Excel doesn't relayout its child widgets (ribbon, XLDESK,
// sheet tabs, status bar) when the embedded Win32 frame is resized
// programmatically — neither SetWindowPos, Application.Width/Height,
// Application.WindowState nor ActiveWindow.WindowState reliably triggers
// it. The initial layout is correct thanks to pre-Open Width/Height in
// LoadExcelFullSta, but subsequent Lister resizes leave Excel's content
// anchored to the original area. Reopening the Lister forces a fresh
// initial layout at the new size.
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
    // We create the brush here and leak it — it lives for the process lifetime.
    wc.hbrBackground = CreateSolidBrush(RGB(200, 200, 200));
    RegisterClassW(&wc);
    registered = true;
}

static void DestroyCloseGuard();   // forward declaration

static void CreateCloseGuard(HINSTANCE hInst)
{
    DestroyCloseGuard();   // idempotent
    if (!g_state.hwndRender) return;

    EnsureGuardClass(hInst);

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
        g_state.hwndRender, nullptr, hInst, nullptr);

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
    LONG_PTR style   = GetWindowLongPtrW(hwndApp, GWL_STYLE);
    LONG_PTR exStyle = GetWindowLongPtrW(hwndApp, GWL_EXSTYLE);

    style &= ~(WS_OVERLAPPED | WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX |
               WS_MAXIMIZEBOX | WS_SYSMENU | WS_DLGFRAME | WS_BORDER | WS_POPUP);
    style |= WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN;
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
    CreateCloseGuard(GetModuleHandleW(nullptr));

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

    VARIANT args[4];
    for (int i = 0; i < 4; ++i) VariantInit(&args[i]);
    args[0].vt = VT_BOOL; args[0].boolVal = VARIANT_FALSE;          // AddToRecentFiles
    args[1].vt = VT_BOOL; args[1].boolVal = VARIANT_TRUE;           // ReadOnly
    args[2].vt = VT_BOOL; args[2].boolVal = VARIANT_FALSE;          // ConfirmConversions
    args[3].vt = VT_BSTR; args[3].bstrVal = SysAllocString(path);   // FileName

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

    VARIANT args[3];
    for (int i = 0; i < 3; ++i) VariantInit(&args[i]);
    args[0].vt = VT_BOOL; args[0].boolVal = VARIANT_TRUE;           // ReadOnly
    args[1].vt = VT_I4;   args[1].lVal    = 0;                      // UpdateLinks = xlUpdateLinksNever
    args[2].vt = VT_BSTR; args[2].bstrVal = SysAllocString(path);   // FileName

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
    g_state.hwndPptApp = nullptr;
}

static void UnloadPptFullSta(bool quitApp)
{
    ClosePptPresentationSta();
    DetachPptWindowSta();
    if (quitApp && g_state.pPptApp)
    {
        DispCall(g_state.pPptApp, L"Quit", DISPATCH_METHOD, nullptr, 0, nullptr);
        g_state.pPptApp->Release();
        g_state.pPptApp = nullptr;
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

    VARIANT args[4];
    for (int i = 0; i < 4; ++i) VariantInit(&args[i]);
    args[0].vt = VT_I4;   args[0].lVal    = -1;                     // WithWindow = msoTrue
    args[1].vt = VT_I4;   args[1].lVal    = 0;                      // Untitled = msoFalse
    args[2].vt = VT_I4;   args[2].lVal    = -1;                     // ReadOnly = msoTrue
    args[3].vt = VT_BSTR; args[3].bstrVal = SysAllocString(path);   // FileName

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
// Preview handler lifecycle. STA thread only.
// ---------------------------------------------------------------------------

static void UnloadHandlerSta()
{
    HideFallbackSta();
    if (g_state.pHandler)
    {
        g_state.pHandler->Unload();
        g_state.pHandler->Release();
        g_state.pHandler = nullptr;
    }
    if (g_state.pHandlerUnk)
    {
        g_state.pHandlerUnk->Release();
        g_state.pHandlerUnk = nullptr;
    }
}

static HRESULT LoadHandlerSta(LPCWSTR path)
{
    UnloadHandlerSta();
    HostLog(L"LoadHandlerSta: path='%s' render=0x%p", path, g_state.hwndRender);

    CLSID clsid;
    HRESULT hr = FindPreviewHandlerClsid(path, &clsid);
    if (FAILED(hr))
    {
        HostLog(L"  FindPreviewHandlerClsid -> 0x%08lX, showing fallback",
                static_cast<long>(hr));
        ShowFallbackSta(path, hr);
        return S_OK;
    }
    {
        wchar_t clsidStr[64] = {};
        StringFromGUID2(clsid, clsidStr, ARRAYSIZE(clsidStr));
        HostLog(L"  CLSID = %s", clsidStr);
    }

    IUnknown* pUnk = nullptr;
    hr = CoCreateInstance(clsid, nullptr,
                          CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER,
                          IID_IUnknown, reinterpret_cast<void**>(&pUnk));
    HostLog(L"  CoCreateInstance -> 0x%08lX", static_cast<long>(hr));
    if (FAILED(hr))
    {
        // Handler is registered but the server cannot be activated — typically
        // a partial / broken Office install. Show the fallback UI rather than
        // an empty pane.
        ShowFallbackSta(path, hr);
        return S_OK;
    }

    bool initialized = false;

    {
        IInitializeWithFile* pInit = nullptr;
        HRESULT qhr = pUnk->QueryInterface(__uuidof(IInitializeWithFile),
                                           reinterpret_cast<void**>(&pInit));
        HostLog(L"  QI IInitializeWithFile -> 0x%08lX", static_cast<long>(qhr));
        if (SUCCEEDED(qhr))
        {
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

    if (!initialized)
    {
        HostLog(L"  no Initialize* interface succeeded — showing fallback");
        pUnk->Release();
        HRESULT reason = FAILED(hr) ? hr : E_NOINTERFACE;
        ShowFallbackSta(path, reason);
        return S_OK;
    }

    IPreviewHandler* pPH = nullptr;
    hr = pUnk->QueryInterface(IID_IPreviewHandler, reinterpret_cast<void**>(&pPH));
    HostLog(L"  QI IPreviewHandler -> 0x%08lX", static_cast<long>(hr));
    if (FAILED(hr))
    {
        pUnk->Release();
        ShowFallbackSta(path, hr);
        return S_OK;
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
        pPH->Release();
        pUnk->Release();
        ShowFallbackSta(path, hr);
        return S_OK;
    }

    g_state.pHandlerUnk = pUnk;
    g_state.pHandler    = pPH;
    HostLog(L"  LoadHandlerSta SUCCESS");
    return S_OK;
}

// ---------------------------------------------------------------------------
// Top-level LOAD dispatcher. Picks quick vs full mode by extension+config and
// hands off to the matching loader. On full-mode failure we degrade silently
// to quick mode so the user always gets *some* preview.
// ---------------------------------------------------------------------------

// Internal worker for LoadFileSta. Mode is supplied explicitly so the
// user-initiated WM_HOST_SWITCH_MODE handler can force the opposite of
// the INI-configured default for the current preview.
static HRESULT LoadFileWithModeSta(LPCWSTR path, AppKind app, Mode mode)
{
    if (g_state.loadingInProgress)
    {
        HostLog(L"LoadFileWithModeSta: ignored — previous load still running");
        return E_ABORT;
    }
    g_state.loadingInProgress = true;

    g_state.currentFile       = path;
    g_state.currentFileApp    = app;
    g_state.currentLoadedMode = mode;
    HostLog(L"LoadFileWithModeSta: app=%d mode=%s",
            static_cast<int>(app),
            mode == Mode::Full ? L"full" : L"quick");

    HRESULT result = E_FAIL;

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
                g_state.loadingInProgress = false;
                return S_OK;
            }
            HostLog(L"  %s full-mode failed (0x%08lX) — falling back to quick",
                    appName, static_cast<long>(hr));
            // Hard-reset all three app slots so quick mode starts from a
            // clean state. Each Unload is a no-op if its slot was empty.
            UnloadWordFullSta(true);
            UnloadExcelFullSta(true);
            UnloadPptFullSta(true);
            // Failed full mode → actual loaded mode will be quick.
            g_state.currentLoadedMode = Mode::Quick;
            // fall through to quick mode
        }
    }

    // Quick mode (or full-mode fallback path). If any Office app's window
    // is still parented under the render pane, detach it first.
    // Detach alone is not enough — we must also close the document so the
    // Office app does not hold a lock on the file. Otherwise a rapid
    // quick→full→quick switch can leave the document open in the background
    // and the next full-mode open may fail or behave unpredictably.
    if (g_state.hwndWordApp)  { CloseWordDocumentSta(); DetachWordWindowSta(); }
    if (g_state.hwndExcelApp) { CloseExcelWorkbookSta(); DetachExcelWindowSta(); }
    if (g_state.hwndPptApp)   { ClosePptPresentationSta(); DetachPptWindowSta(); }

    result = LoadHandlerSta(path);
    UpdateModeButtonSta();
    // Quick-mode preview handlers may also do async window layout; a
    // shorter delay is enough since they don't need Office startup time.
    if (g_state.hwndRender)
        SetTimer(g_state.hwndRender, kModeButtonZTimerId, 200, nullptr);
    g_state.loadingInProgress = false;
    return result;
}

// Public LoadFile entry point — uses the INI-configured mode. This is the
// path taken by the WM_HOST_LOAD pipe command (one per file shown in the
// Lister), so switching files always resets to the configured default and
// the user's button click does not carry over.
static HRESULT LoadFileSta(LPCWSTR path)
{
    AppKind app  = ClassifyByExtension(path);
    Mode    mode = SelectMode(app);
    return LoadFileWithModeSta(path, app, mode);
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
            HRESULT hr = LoadFileWithModeSta(path.c_str(), g_state.currentFileApp, target);
            if (FAILED(hr))
            {
                HostLog(L"WM_HOST_SWITCH_MODE: LoadFileWithModeSta failed hr=0x%08lX",
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
    bool gracefulClose = false;

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
            // STA's WM_HOST_CLOSE handler writes "OK" and posts WM_QUIT.
            PostMessageW(g_state.hwndSta, WM_HOST_CLOSE, 0, 0);
            gracefulClose = true;
            break;
        }
        else
        {
            PipeWriteUtf16(g_state.hPipe, L"ERR unknown command\n");
        }
    }

    delete[] buf;

    // If the pipe died unexpectedly, kick the STA out of its message loop.
    if (!gracefulClose)
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

    HRESULT hr = CoInitializeEx(nullptr,
                                COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    HostLog(L"  CoInitializeEx -> 0x%08lX", static_cast<long>(hr));
    if (FAILED(hr))
        return 2;

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
