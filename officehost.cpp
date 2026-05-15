/*
 * officehost.cpp - Total Commander Office preview host process.
 *
 * Spawned by tcoffice.wlx. Hosts a Windows Preview Handler COM object in a
 * child HWND owned by Total Commander and receives commands from the plugin
 * DLL over a named pipe.
 *
 * Threading model:
 *   - Main thread is the STA. Runs a Win32 message loop and a hidden
 *     message-only window. All COM calls happen on this thread.
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

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <windows.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <objbase.h>
#include <stdio.h>
#include <string>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "user32.lib")

// ---------------------------------------------------------------------------
// COM interface declarations
//
// IInitializeWithFile / IInitializeWithStream live in propsys.h, but we only
// need their IID and a minimal vtable. Declaring them by hand keeps us off
// propsys.lib at link time.
// ---------------------------------------------------------------------------

struct __declspec(uuid("b7d14566-0509-4cce-a71f-0a554233bd9b")) IInitializeWithFile
    : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Initialize(LPCWSTR pszFilePath, DWORD grfMode) = 0;
};

struct __declspec(uuid("b824b49d-22ac-4132-ac65-c0a17b6a3b9c")) IInitializeWithStream
    : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Initialize(IStream* pstream, DWORD grfMode) = 0;
};

// Category GUID for Windows Preview Handlers — the subkey
// HKCR\.<ext>\shellex\{8895b1c6-...} points to the handler's CLSID.
static const GUID kPreviewHandlerCategory =
    { 0x8895b1c6, 0xb41f, 0x4c1c, { 0xa5, 0x62, 0x0d, 0x56, 0x42, 0x50, 0x83, 0x6f } };

// ---------------------------------------------------------------------------
// Process-wide state. Modified only from the STA thread.
// ---------------------------------------------------------------------------

struct HostState
{
    HWND                hwndParent  = nullptr;            // child window provided by the plugin
    HWND                hwndSta     = nullptr;            // message-only window for cross-thread dispatch
    HANDLE              hPipe       = INVALID_HANDLE_VALUE;
    IUnknown*           pHandlerUnk = nullptr;            // raw COM object
    IPreviewHandler*    pHandler    = nullptr;            // queried IPreviewHandler view
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

    // 1) HKCR\.<ext>\shellex\{cat}
    std::wstring sub = ext + L"\\shellex\\" + catGuid;
    std::wstring clsidStr;
    if (!ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
    {
        // 2) Resolve ProgID then try HKCR\<ProgID>\shellex\{cat}
        std::wstring progId;
        if (!ReadDefaultString(HKEY_CLASSES_ROOT, ext.c_str(), progId) || progId.empty())
            return REGDB_E_CLASSNOTREG;
        sub = progId + L"\\shellex\\" + catGuid;
        if (!ReadDefaultString(HKEY_CLASSES_ROOT, sub.c_str(), clsidStr))
            return REGDB_E_CLASSNOTREG;
    }
    return CLSIDFromString(clsidStr.c_str(), out);
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

    CLSID clsid;
    HRESULT hr = FindPreviewHandlerClsid(path, &clsid);
    if (FAILED(hr)) return hr;

    IUnknown* pUnk = nullptr;
    hr = CoCreateInstance(clsid, nullptr,
                          CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER,
                          IID_IUnknown, reinterpret_cast<void**>(&pUnk));
    if (FAILED(hr)) return hr;

    bool initialized = false;

    // Prefer IInitializeWithFile (Office handlers support it directly).
    {
        IInitializeWithFile* pInit = nullptr;
        if (SUCCEEDED(pUnk->QueryInterface(__uuidof(IInitializeWithFile),
                                           reinterpret_cast<void**>(&pInit))))
        {
            HRESULT ihr = pInit->Initialize(path, STGM_READ);
            pInit->Release();
            initialized = SUCCEEDED(ihr);
            if (!initialized) hr = ihr;
        }
    }

    // Fallback: IInitializeWithStream via SHCreateStreamOnFileEx.
    if (!initialized)
    {
        IInitializeWithStream* pInit = nullptr;
        if (SUCCEEDED(pUnk->QueryInterface(__uuidof(IInitializeWithStream),
                                           reinterpret_cast<void**>(&pInit))))
        {
            IStream* pStream = nullptr;
            HRESULT shr = SHCreateStreamOnFileEx(
                path, STGM_READ | STGM_SHARE_DENY_NONE,
                0, FALSE, nullptr, &pStream);
            if (SUCCEEDED(shr))
            {
                HRESULT ihr = pInit->Initialize(pStream, STGM_READ);
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
        pUnk->Release();
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    IPreviewHandler* pPH = nullptr;
    hr = pUnk->QueryInterface(IID_IPreviewHandler, reinterpret_cast<void**>(&pPH));
    if (FAILED(hr))
    {
        pUnk->Release();
        return hr;
    }

    RECT rc = {};
    GetClientRect(g_state.hwndParent, &rc);
    hr = pPH->SetWindow(g_state.hwndParent, &rc);
    if (SUCCEEDED(hr))
        hr = pPH->DoPreview();

    if (FAILED(hr))
    {
        pPH->Release();
        pUnk->Release();
        return hr;
    }

    g_state.pHandlerUnk = pUnk;
    g_state.pHandler    = pPH;
    return S_OK;
}

static void ResizeHandlerSta(int w, int h)
{
    if (!g_state.pHandler) return;
    RECT rc = { 0, 0, w, h };
    g_state.pHandler->SetRect(&rc);
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
            UnloadHandlerSta();
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
    wc.lpfnWndProc = StaWndProc;
    wc.hInstance = hInst;
    wc.lpszClassName = kClass;
    RegisterClassW(&wc);
    return CreateWindowExW(0, kClass, L"", 0, 0, 0, 0, 0,
                           HWND_MESSAGE, nullptr, hInst, nullptr);
}

int wmain(int argc, wchar_t** argv)
{
    std::wstring pipeName;
    if (!ParseArgs(argc, argv, &g_state.hwndParent, &pipeName))
    {
        OutputDebugStringW(L"officehost: usage --hwnd <handle> --pipe <name>\n");
        return 1;
    }

    HRESULT hr = CoInitializeEx(nullptr,
                                COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr))
    {
        OutputDebugStringW(L"officehost: CoInitializeEx failed\n");
        return 2;
    }

    // Pipe was created by the plugin with FILE_FLAG_OVERLAPPED. Open the
    // client end with the matching flag.
    g_state.hPipe = CreateFileW(pipeName.c_str(),
                                GENERIC_READ | GENERIC_WRITE,
                                0, nullptr, OPEN_EXISTING,
                                FILE_FLAG_OVERLAPPED, nullptr);
    if (g_state.hPipe == INVALID_HANDLE_VALUE)
    {
        OutputDebugStringW(L"officehost: CreateFile on pipe failed\n");
        CoUninitialize();
        return 3;
    }
    DWORD mode = PIPE_READMODE_MESSAGE;
    SetNamedPipeHandleState(g_state.hPipe, &mode, nullptr, nullptr);

    g_state.hwndSta = CreateStaWindow(GetModuleHandleW(nullptr));
    if (!g_state.hwndSta)
    {
        CloseHandle(g_state.hPipe);
        CoUninitialize();
        return 4;
    }

    HANDLE hThread = CreateThread(nullptr, 0, PipeReaderThread, nullptr, 0, nullptr);
    if (!hThread)
    {
        DestroyWindow(g_state.hwndSta);
        CloseHandle(g_state.hPipe);
        CoUninitialize();
        return 5;
    }

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    // Make sure the reader exits before we close the pipe handle it shares.
    // CloseHandle unblocks any pending PipeReadSync with ERROR_OPERATION_ABORTED.
    HANDLE hPipe = g_state.hPipe;
    g_state.hPipe = INVALID_HANDLE_VALUE;
    CancelIoEx(hPipe, nullptr);
    WaitForSingleObject(hThread, 2000);
    CloseHandle(hThread);
    CloseHandle(hPipe);

    DestroyWindow(g_state.hwndSta);
    CoUninitialize();
    return 0;
}
