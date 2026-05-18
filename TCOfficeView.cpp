// SPDX-License-Identifier: Apache-2.0
// Copyright 2026 Michal Růžička <ruzicka.mich@gmail.com>

/*
 * TCOfficeView.cpp - Total Commander Lister plugin for MS Office documents.
 *
 * Flow:
 *   1. TC calls ListLoad / ListLoadW with a file path and the parent HWND.
 *   2. The plugin creates a child window inside that parent (rendering target).
 *   3. It creates a named pipe (server side, overlapped) and spawns
 *      TCOfficeViewHost.exe with --hwnd <child> --pipe <name>.
 *   4. ConnectNamedPipe waits with a 5-second timeout; if the host process
 *      exits before connecting, the wait returns early.
 *   5. Once connected, the plugin sends "LOAD <path>". The host CoCreateInstances
 *      the registered IPreviewHandler for the extension and renders into the
 *      child HWND.
 *   6. On WM_SIZE the plugin coalesces resize events (50 ms timer) and
 *      forwards them as RESIZE <w> <h>.
 *   7. ListCloseWindow sends CLOSE and waits briefly for the host to exit.
 */

#include "listplug.h"
#include <windows.h>
#include <string>
#include <map>
#include <mutex>
#include <sstream>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "advapi32.lib")

// ---------------------------------------------------------------------------
// Global state
// ---------------------------------------------------------------------------

struct PreviewSession
{
    HWND            hwndChild   = nullptr;      // rendering target
    HANDLE          hProcess    = nullptr;      // TCOfficeViewHost.exe
    HANDLE          hPipe       = INVALID_HANDLE_VALUE;
    std::wstring    pipeName;
    std::wstring    currentFile;
};

static std::map<HWND, PreviewSession*> g_sessions;
static std::mutex                       g_sessionsMutex;
static HINSTANCE                        g_hInstance = nullptr;

static const wchar_t* kWndClassName  = L"TCOfficeViewPreview";
static const UINT_PTR kResizeTimerId = 1;

// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------

static std::wstring GetPluginDir()
{
    wchar_t path[MAX_PATH] = {};
    GetModuleFileNameW(g_hInstance, path, MAX_PATH);
    std::wstring s(path);
    auto pos = s.find_last_of(L"\\/");
    return (pos != std::wstring::npos) ? s.substr(0, pos) : L".";
}

static std::wstring MakePipeName()
{
    // \\.\pipe\TCOfficeView_<pid>_<tick>
    wchar_t buf[128] = {};
    swprintf_s(buf, L"\\\\.\\pipe\\TCOfficeView_%lu_%llu",
               GetCurrentProcessId(),
               static_cast<unsigned long long>(GetTickCount64()));
    return buf;
}

static std::wstring Utf8ToWide(const char* utf8)
{
    if (!utf8) return L"";
    int n = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, nullptr, 0);
    if (n <= 1) return L"";
    std::wstring w(static_cast<size_t>(n - 1), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, &w[0], n);
    return w;
}

// ---------------------------------------------------------------------------
// Pipe IO. The pipe is opened with FILE_FLAG_OVERLAPPED so we can apply a
// timeout to ConnectNamedPipe. Subsequent reads/writes must therefore supply
// an OVERLAPPED structure even when we want them to appear synchronous.
// ---------------------------------------------------------------------------

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

static BOOL PipeWriteUtf16(HANDLE h, const std::wstring& text)
{
    return PipeWriteSync(h, text.data(),
                         static_cast<DWORD>(text.size() * sizeof(wchar_t)));
}

// Wait up to timeoutMs for the host process to connect to the pipe. If the
// host process dies during the wait, returns early.
static bool ConnectWithTimeout(HANDLE hPipe, HANDLE hProcess, DWORD timeoutMs)
{
    OVERLAPPED ov = {};
    ov.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);

    BOOL connected = ConnectNamedPipe(hPipe, &ov);
    DWORD err = GetLastError();

    bool done = false;
    if (connected)
    {
        done = true;                                 // unusual but legal
    }
    else if (err == ERROR_PIPE_CONNECTED)
    {
        done = true;                                 // client got there first
    }
    else if (err == ERROR_IO_PENDING)
    {
        HANDLE waits[2] = { ov.hEvent, hProcess };
        DWORD r = WaitForMultipleObjects(2, waits, FALSE, timeoutMs);
        if (r == WAIT_OBJECT_0)
        {
            done = true;
        }
        else
        {
            // Either the host died (WAIT_OBJECT_0 + 1) or we timed out.
            CancelIoEx(hPipe, &ov);
            WaitForSingleObject(ov.hEvent, 200);     // drain cancellation
        }
    }

    CloseHandle(ov.hEvent);
    return done;
}

// ---------------------------------------------------------------------------
// Child window procedure
// ---------------------------------------------------------------------------

static LRESULT CALLBACK ChildWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_SIZE:
        {
            // Coalesce: WM_SIZE fires per-pixel during a drag. We only want
            // to write to the pipe at most ~20 Hz, otherwise the resize loop
            // stalls on pipe IO.
            SetTimer(hWnd, kResizeTimerId, 50, nullptr);
            return 0;
        }
        case WM_TIMER:
        {
            if (wp != kResizeTimerId) break;
            KillTimer(hWnd, kResizeTimerId);

            RECT rc = {};
            GetClientRect(hWnd, &rc);
            const int w = rc.right - rc.left;
            const int h = rc.bottom - rc.top;

            PreviewSession* session = nullptr;
            {
                std::lock_guard<std::mutex> lock(g_sessionsMutex);
                auto it = g_sessions.find(hWnd);
                if (it != g_sessions.end()) session = it->second;
            }
            if (session && session->hPipe != INVALID_HANDLE_VALUE)
            {
                wchar_t cmd[64] = {};
                int n = swprintf_s(cmd, L"RESIZE %d %d\n", w, h);
                if (n > 0)
                    PipeWriteSync(session->hPipe, cmd,
                                  static_cast<DWORD>(n * sizeof(wchar_t)));
            }
            return 0;
        }
        case WM_ERASEBKGND:
            // The preview handler paints the whole window — skip background
            // erase to avoid flicker.
            return 1;
    }
    return DefWindowProcW(hWnd, msg, wp, lp);
}

static void EnsureWindowClass()
{
    static bool registered = false;
    if (registered) return;

    WNDCLASSW wc = {};
    wc.lpfnWndProc   = ChildWndProc;
    wc.hInstance     = g_hInstance;
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.lpszClassName = kWndClassName;
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    RegisterClassW(&wc);
    registered = true;
}

// ---------------------------------------------------------------------------
// Host process launch
// ---------------------------------------------------------------------------

static bool LaunchHost(PreviewSession* session, const std::wstring& file)
{
    session->pipeName = MakePipeName();
    session->hPipe = CreateNamedPipeW(
        session->pipeName.c_str(),
        PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED,
        PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
        1, 64 * 1024, 64 * 1024, 0, nullptr);
    if (session->hPipe == INVALID_HANDLE_VALUE) return false;

    // Each bitness of the plugin ships its own host EXE; both files live in
    // the same TC plugin folder, so they must have distinct names.
#ifdef _WIN64
    std::wstring exePath = GetPluginDir() + L"\\TCOfficeViewHost.exe";
#else
    std::wstring exePath = GetPluginDir() + L"\\TCOfficeViewHost_x86.exe";
#endif
    std::wstringstream cmdLineStream;
    cmdLineStream << L"\"" << exePath << L"\""
                  << L" --hwnd " << reinterpret_cast<uintptr_t>(session->hwndChild)
                  << L" --pipe \"" << session->pipeName << L"\"";
    std::wstring cmdLine = cmdLineStream.str();

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = {};
    BOOL ok = CreateProcessW(
        exePath.c_str(),
        cmdLine.data(),
        nullptr, nullptr, FALSE,
        CREATE_NO_WINDOW,
        nullptr, nullptr, &si, &pi);
    if (!ok)
    {
        CloseHandle(session->hPipe);
        session->hPipe = INVALID_HANDLE_VALUE;
        return false;
    }
    session->hProcess = pi.hProcess;
    CloseHandle(pi.hThread);

    if (!ConnectWithTimeout(session->hPipe, session->hProcess, 5000))
    {
        TerminateProcess(session->hProcess, 1);
        CloseHandle(session->hProcess);
        session->hProcess = nullptr;
        CloseHandle(session->hPipe);
        session->hPipe = INVALID_HANDLE_VALUE;
        return false;
    }

    std::wstring loadCmd = L"LOAD " + file + L"\n";
    PipeWriteUtf16(session->hPipe, loadCmd);
    return true;
}

static void TerminateSession(PreviewSession* session)
{
    if (!session) return;

    if (session->hPipe != INVALID_HANDLE_VALUE)
    {
        PipeWriteUtf16(session->hPipe, L"CLOSE\n");
        if (session->hProcess)
        {
            if (WaitForSingleObject(session->hProcess, 2000) != WAIT_OBJECT_0)
                TerminateProcess(session->hProcess, 1);
            CloseHandle(session->hProcess);
        }
        CloseHandle(session->hPipe);
    }
    delete session;
}

// ---------------------------------------------------------------------------
// Exported TC plugin API
// ---------------------------------------------------------------------------

HWND __stdcall ListLoadW(HWND ParentWin, wchar_t* FileToLoad, int /*ShowFlags*/)
{
    EnsureWindowClass();

    RECT rc = {};
    GetClientRect(ParentWin, &rc);

    HWND hChild = CreateWindowExW(
        0, kWndClassName, L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
        0, 0, rc.right, rc.bottom,
        ParentWin, nullptr, g_hInstance, nullptr);
    if (!hChild) return nullptr;

    auto* session = new PreviewSession();
    session->hwndChild   = hChild;
    session->currentFile = FileToLoad;

    {
        std::lock_guard<std::mutex> lock(g_sessionsMutex);
        g_sessions[hChild] = session;
    }

    if (!LaunchHost(session, FileToLoad))
    {
        // Fall back to TC's built-in viewer by returning nullptr.
        {
            std::lock_guard<std::mutex> lock(g_sessionsMutex);
            g_sessions.erase(hChild);
        }
        DestroyWindow(hChild);
        delete session;
        return nullptr;
    }

    return hChild;
}

HWND __stdcall ListLoad(HWND ParentWin, char* FileToLoad, int ShowFlags)
{
    std::wstring wide = Utf8ToWide(FileToLoad);
    return ListLoadW(ParentWin, wide.data(), ShowFlags);
}

int __stdcall ListLoadNextW(HWND /*ParentWin*/, HWND PluginWin,
                            wchar_t* FileToLoad, int /*ShowFlags*/)
{
    PreviewSession* session = nullptr;
    {
        std::lock_guard<std::mutex> lock(g_sessionsMutex);
        auto it = g_sessions.find(PluginWin);
        if (it != g_sessions.end()) session = it->second;
    }
    if (!session || session->hPipe == INVALID_HANDLE_VALUE) return 0;

    std::wstring loadCmd = std::wstring(L"LOAD ") + FileToLoad + L"\n";
    BOOL ok = PipeWriteUtf16(session->hPipe, loadCmd);
    if (ok) session->currentFile = FileToLoad;
    return ok ? 1 : 0;
}

int __stdcall ListLoadNext(HWND ParentWin, HWND PluginWin,
                           char* FileToLoad, int ShowFlags)
{
    std::wstring wide = Utf8ToWide(FileToLoad);
    return ListLoadNextW(ParentWin, PluginWin, wide.data(), ShowFlags);
}

void __stdcall ListCloseWindow(HWND ListWin)
{
    PreviewSession* session = nullptr;
    {
        std::lock_guard<std::mutex> lock(g_sessionsMutex);
        auto it = g_sessions.find(ListWin);
        if (it != g_sessions.end())
        {
            session = it->second;
            g_sessions.erase(it);
        }
    }
    TerminateSession(session);
    DestroyWindow(ListWin);
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen)
{
    const char* s = "EXT=\"DOCX\"|EXT=\"DOC\"|EXT=\"DOCM\"|"
                    "EXT=\"XLSX\"|EXT=\"XLS\"|EXT=\"XLSM\"|EXT=\"XLSB\"|"
                    "EXT=\"PPTX\"|EXT=\"PPT\"|EXT=\"PPTM\"|"
                    "EXT=\"RTF\"|EXT=\"VSDX\"|EXT=\"MSG\"";
    strncpy_s(DetectString, maxlen, s, _TRUNCATE);
}

int __stdcall ListSendCommand(HWND /*ListWin*/, int /*Command*/, int /*Parameter*/)
{
    // Zoom / print / find / copy not yet wired through to the host.
    return 0;
}

// ---------------------------------------------------------------------------
// DllMain
// ---------------------------------------------------------------------------

BOOL APIENTRY DllMain(HINSTANCE hInst, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hInstance = hInst;
        DisableThreadLibraryCalls(hInst);
    }
    else if (reason == DLL_PROCESS_DETACH)
    {
        // Belt-and-braces: if TC forgot to close a session, tear it down.
        std::lock_guard<std::mutex> lock(g_sessionsMutex);
        for (auto& kv : g_sessions) TerminateSession(kv.second);
        g_sessions.clear();
    }
    return TRUE;
}
