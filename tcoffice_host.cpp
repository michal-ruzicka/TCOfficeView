/*
 * tcoffice_host.cpp - Total Commander Office preview host process.
 *
 * Spawned by tcoffice.wlx. Hosts a Windows Preview Handler COM object and
 * renders it inside Total Commander's Lister pane.
 *
 * Window topology:
 *
 *      Total Commander process              this host process
 *      ---------------------                ------------------
 *      Lister parent
 *        └── plugin child   (tcoffice DLL)
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
#include <propsys.h>          // IInitializeWithFile, IInitializeWithStream
#include <objbase.h>
#include <stdio.h>
#include <string>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "user32.lib")

// ---------------------------------------------------------------------------
// Diagnostic logging
//
// Writes to %TEMP%\tcoffice_host.log. Cheap and synchronous so we can trace
// the exact step that fails when a preview doesn't render. Set
// HOST_LOG_ENABLED to 0 to silence.
// ---------------------------------------------------------------------------

#define HOST_LOG_ENABLED 1

#if HOST_LOG_ENABLED
static void HostLog(const wchar_t* fmt, ...)
{
    static HANDLE hLog = nullptr;
    static CRITICAL_SECTION cs;
    static bool csInit = false;
    if (!csInit) { InitializeCriticalSection(&cs); csInit = true; }
    EnterCriticalSection(&cs);
    if (!hLog)
    {
        wchar_t path[MAX_PATH] = {};
        GetTempPathW(MAX_PATH, path);
        wcscat_s(path, L"tcoffice_host.log");
        hLog = CreateFileW(path, FILE_APPEND_DATA,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    }
    if (hLog && hLog != INVALID_HANDLE_VALUE)
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
        WriteFile(hLog, prefix,
                  static_cast<DWORD>(wcslen(prefix) * sizeof(wchar_t)),
                  &written, nullptr);
        WriteFile(hLog, body,
                  static_cast<DWORD>(wcslen(body) * sizeof(wchar_t)),
                  &written, nullptr);
        static const wchar_t nl[] = L"\r\n";
        WriteFile(hLog, nl, sizeof(nl) - sizeof(wchar_t), &written, nullptr);
        FlushFileBuffers(hLog);
    }
    LeaveCriticalSection(&cs);
}
#else
static void HostLog(const wchar_t*, ...) {}
#endif

// Category GUID for Windows Preview Handlers — the subkey
// HKCR\.<ext>\shellex\{8895b1c6-...} points to the handler's CLSID.
static const GUID kPreviewHandlerCategory =
    { 0x8895b1c6, 0xb41f, 0x4c1c, { 0xa5, 0x62, 0x0d, 0x56, 0x42, 0x50, 0x83, 0x6f } };

// ---------------------------------------------------------------------------
// Process-wide state. Modified only from the STA thread.
// ---------------------------------------------------------------------------

struct HostState
{
    HWND                hwndPluginChild = nullptr;        // window inside TC, supplied on cmd line
    HWND                hwndRender      = nullptr;        // our own child, reparented into hwndPluginChild
    HWND                hwndSta         = nullptr;        // message-only window for cross-thread dispatch
    HANDLE              hPipe           = INVALID_HANDLE_VALUE;
    IUnknown*           pHandlerUnk     = nullptr;        // raw COM object
    IPreviewHandler*    pHandler        = nullptr;        // queried IPreviewHandler view
};

static HostState g_state;

// Message IDs used by the worker thread to call into the STA thread.
// Each handler returns 0 on success, or an HRESULT cast to LRESULT on error.
static constexpr UINT WM_HOST_LOAD   = WM_USER + 1;
static constexpr UINT WM_HOST_RESIZE = WM_USER + 2;
static constexpr UINT WM_HOST_CLOSE  = WM_USER + 3;

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

static const wchar_t* kRenderClassName = L"TCOfficeHostRender";

static LRESULT CALLBACK RenderWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (msg == WM_ERASEBKGND) return 1;        // preview handler paints
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
    return true;
}

static void DestroyRenderWindowSta()
{
    if (g_state.hwndRender)
    {
        // Reparent back to HWND_MESSAGE before destroying — this severs the
        // cross-process link cleanly so TC's UI thread doesn't see a stale
        // child reference while we're tearing down.
        SetParent(g_state.hwndRender, HWND_MESSAGE);
        DestroyWindow(g_state.hwndRender);
        g_state.hwndRender = nullptr;
    }
}

// ---------------------------------------------------------------------------
// Preview handler lifecycle. STA thread only.
// ---------------------------------------------------------------------------

static void UnloadHandlerSta()
{
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
        HostLog(L"  FindPreviewHandlerClsid -> 0x%08lX", static_cast<long>(hr));
        return hr;
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
    if (FAILED(hr)) return hr;

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
        HostLog(L"  no Initialize* interface succeeded — bailing");
        pUnk->Release();
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    IPreviewHandler* pPH = nullptr;
    hr = pUnk->QueryInterface(IID_IPreviewHandler, reinterpret_cast<void**>(&pPH));
    HostLog(L"  QI IPreviewHandler -> 0x%08lX", static_cast<long>(hr));
    if (FAILED(hr))
    {
        pUnk->Release();
        return hr;
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
        return hr;
    }

    g_state.pHandlerUnk = pUnk;
    g_state.pHandler    = pPH;
    HostLog(L"  LoadHandlerSta SUCCESS");
    return S_OK;
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
}

// ---------------------------------------------------------------------------
// STA window procedure
// ---------------------------------------------------------------------------

static LRESULT CALLBACK StaWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_HOST_LOAD:
        {
            LPCWSTR path = reinterpret_cast<LPCWSTR>(lp);
            HRESULT hr = LoadHandlerSta(path);
            return SUCCEEDED(hr) ? 0 : static_cast<LRESULT>(static_cast<long>(hr));
        }
        case WM_HOST_RESIZE:
        {
            int w = static_cast<int>(LOWORD(wp));
            int h = static_cast<int>(HIWORD(wp));
            ResizeHandlerSta(w, h);
            return 0;
        }
        case WM_HOST_CLOSE:
        {
            HostLog(L"WM_HOST_CLOSE");
            UnloadHandlerSta();
            DestroyRenderWindowSta();
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
            LRESULT r = SendMessageW(g_state.hwndSta, WM_HOST_LOAD,
                                     0, reinterpret_cast<LPARAM>(buf + 5));
            if (r == 0)
            {
                PipeWriteUtf16(g_state.hPipe, L"OK\n");
            }
            else
            {
                wchar_t errMsg[96] = {};
                _snwprintf_s(errMsg, ARRAYSIZE(errMsg), _TRUNCATE,
                             L"ERR LoadHandler hr=0x%08lX\n",
                             static_cast<long>(r));
                PipeWriteUtf16(g_state.hPipe, errMsg);
            }
        }
        else if (wcsncmp(buf, L"RESIZE ", 7) == 0)
        {
            int w = 0, h = 0;
            if (swscanf_s(buf + 7, L"%d %d", &w, &h) == 2 && w > 0 && h > 0)
            {
                WPARAM wp = MAKELONG(w, h);
                SendMessageW(g_state.hwndSta, WM_HOST_RESIZE, wp, 0);
            }
            // RESIZE is fire-and-forget — no response.
        }
        else if (wcsncmp(buf, L"CLOSE", 5) == 0)
        {
            SendMessageW(g_state.hwndSta, WM_HOST_CLOSE, 0, 0);
            PipeWriteUtf16(g_state.hPipe, L"OK\n");
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
    static const wchar_t* kClass = L"TCOfficeHostSta";
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

    CoUninitialize();
    HostLog(L"wmain: exit");
    return 0;
}
