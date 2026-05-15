/*
 * tcoffice.cpp - Total Commander Lister plugin pro MS Office dokumenty
 *
 * Architektura:
 *   1. TC zavolá ListLoad/ListLoadW s cestou k souboru a HWND parent okna
 *   2. Plugin vytvoří child okno (placeholder pro preview handler)
 *   3. Spustí se OfficePreviewHost.exe s parametry (HWND, pipe name)
 *   4. Plugin posílá příkazy přes named pipe (LOAD, RESIZE, CLOSE)
 *   5. Host process si v dceřiném okně vytvoří Office preview handler
 *      přes COM (IPreviewHandler) a renderuje přímo do okna
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
// Globální stav
// ---------------------------------------------------------------------------

struct PreviewSession {
    HWND        hwndChild;       // child window (rendering target)
    HANDLE      hProcess;        // OfficePreviewHost.exe
    HANDLE      hPipe;           // named pipe server handle
    std::wstring pipeName;
    std::wstring currentFile;
};

static std::map<HWND, PreviewSession*> g_sessions;
static std::mutex g_sessionsMutex;
static HINSTANCE g_hInstance = nullptr;

static const wchar_t* WND_CLASS_NAME = L"TCOfficePreviewHost";

// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------

static std::wstring GetPluginDir() {
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(g_hInstance, path, MAX_PATH);
    std::wstring s(path);
    auto pos = s.find_last_of(L"\\/");
    return (pos != std::wstring::npos) ? s.substr(0, pos) : L".";
}

static std::wstring MakePipeName() {
    // Unikátní jméno per session: \\.\pipe\tcoffice_<pid>_<tick>
    wchar_t buf[128];
    swprintf_s(buf, L"\\\\.\\pipe\\tcoffice_%lu_%llu",
               GetCurrentProcessId(),
               (unsigned long long)GetTickCount64());
    return buf;
}

static std::wstring Utf8ToWide(const char* utf8) {
    if (!utf8) return L"";
    int n = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, nullptr, 0);
    std::wstring w(n - 1, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, &w[0], n);
    return w;
}

// ---------------------------------------------------------------------------
// Child window procedure
// ---------------------------------------------------------------------------

static LRESULT CALLBACK ChildWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_SIZE: {
            // Předej resize do host procesu přes pipe
            std::lock_guard<std::mutex> lock(g_sessionsMutex);
            auto it = g_sessions.find(hWnd);
            if (it != g_sessions.end() && it->second->hPipe != INVALID_HANDLE_VALUE) {
                wchar_t cmd[64];
                swprintf_s(cmd, L"RESIZE %d %d\n", LOWORD(lp), HIWORD(lp));
                DWORD written;
                WriteFile(it->second->hPipe, cmd,
                         (DWORD)(wcslen(cmd) * sizeof(wchar_t)), &written, nullptr);
            }
            return 0;
        }
        case WM_ERASEBKGND:
            // Host process renderuje celé okno, vyhneme se blikání
            return 1;
    }
    return DefWindowProcW(hWnd, msg, wp, lp);
}

static void EnsureWindowClass() {
    static bool registered = false;
    if (registered) return;

    WNDCLASSW wc = {};
    wc.lpfnWndProc = ChildWndProc;
    wc.hInstance = g_hInstance;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = WND_CLASS_NAME;
    wc.style = CS_HREDRAW | CS_VREDRAW;
    RegisterClassW(&wc);
    registered = true;
}

// ---------------------------------------------------------------------------
// Spuštění host procesu
// ---------------------------------------------------------------------------

static bool LaunchHost(PreviewSession* session, const std::wstring& file) {
    // Vytvoř pipe server
    session->pipeName = MakePipeName();
    session->hPipe = CreateNamedPipeW(
        session->pipeName.c_str(),
        PIPE_ACCESS_DUPLEX,
        PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
        1, 4096, 4096, 0, nullptr);
    if (session->hPipe == INVALID_HANDLE_VALUE) return false;

    // Spusť OfficePreviewHost.exe
    std::wstring exePath = GetPluginDir() + L"\\OfficePreviewHost.exe";
    std::wstringstream cmdLine;
    cmdLine << L"\"" << exePath << L"\""
            << L" --hwnd " << (uintptr_t)session->hwndChild
            << L" --pipe \"" << session->pipeName << L"\"";

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = {};
    std::wstring cmd = cmdLine.str();

    BOOL ok = CreateProcessW(
        exePath.c_str(),
        &cmd[0],
        nullptr, nullptr, FALSE,
        CREATE_NO_WINDOW,
        nullptr, nullptr, &si, &pi);

    if (!ok) {
        CloseHandle(session->hPipe);
        session->hPipe = INVALID_HANDLE_VALUE;
        return false;
    }

    session->hProcess = pi.hProcess;
    CloseHandle(pi.hThread);

    // Čekej až se host připojí k pipe (s timeoutem)
    OVERLAPPED ov = {};
    ov.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    ConnectNamedPipe(session->hPipe, &ov);
    DWORD waitResult = WaitForSingleObject(ov.hEvent, 5000);
    CloseHandle(ov.hEvent);

    if (waitResult != WAIT_OBJECT_0) {
        TerminateProcess(session->hProcess, 1);
        CloseHandle(session->hProcess);
        CloseHandle(session->hPipe);
        return false;
    }

    // Pošli LOAD příkaz
    std::wstring loadCmd = L"LOAD " + file + L"\n";
    DWORD written;
    WriteFile(session->hPipe, loadCmd.c_str(),
             (DWORD)(loadCmd.size() * sizeof(wchar_t)), &written, nullptr);

    return true;
}

static void TerminateSession(PreviewSession* session) {
    if (!session) return;

    if (session->hPipe != INVALID_HANDLE_VALUE) {
        const wchar_t* closeCmd = L"CLOSE\n";
        DWORD written;
        WriteFile(session->hPipe, closeCmd,
                 (DWORD)(wcslen(closeCmd) * sizeof(wchar_t)), &written, nullptr);

        // Počkej krátce na čistý exit, jinak zabij
        if (session->hProcess) {
            if (WaitForSingleObject(session->hProcess, 2000) != WAIT_OBJECT_0) {
                TerminateProcess(session->hProcess, 1);
            }
            CloseHandle(session->hProcess);
        }
        CloseHandle(session->hPipe);
    }
    delete session;
}

// ---------------------------------------------------------------------------
// Exportované funkce TC API
// ---------------------------------------------------------------------------

HWND __stdcall ListLoadW(HWND ParentWin, wchar_t* FileToLoad, int ShowFlags) {
    EnsureWindowClass();

    RECT rc;
    GetClientRect(ParentWin, &rc);

    HWND hChild = CreateWindowExW(
        0, WND_CLASS_NAME, L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
        0, 0, rc.right, rc.bottom,
        ParentWin, nullptr, g_hInstance, nullptr);

    if (!hChild) return nullptr;

    auto* session = new PreviewSession();
    session->hwndChild = hChild;
    session->hProcess = nullptr;
    session->hPipe = INVALID_HANDLE_VALUE;
    session->currentFile = FileToLoad;

    {
        std::lock_guard<std::mutex> lock(g_sessionsMutex);
        g_sessions[hChild] = session;
    }

    if (!LaunchHost(session, FileToLoad)) {
        // Nepodařilo se - vrátíme NULL, TC použije svůj interní viewer
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

HWND __stdcall ListLoad(HWND ParentWin, char* FileToLoad, int ShowFlags) {
    std::wstring wide = Utf8ToWide(FileToLoad);
    return ListLoadW(ParentWin, &wide[0], ShowFlags);
}

int __stdcall ListLoadNextW(HWND ParentWin, HWND PluginWin,
                            wchar_t* FileToLoad, int ShowFlags) {
    std::lock_guard<std::mutex> lock(g_sessionsMutex);
    auto it = g_sessions.find(PluginWin);
    if (it == g_sessions.end()) return 0;

    PreviewSession* session = it->second;
    if (session->hPipe == INVALID_HANDLE_VALUE) return 0;

    std::wstring loadCmd = std::wstring(L"LOAD ") + FileToLoad + L"\n";
    DWORD written;
    BOOL ok = WriteFile(session->hPipe, loadCmd.c_str(),
                       (DWORD)(loadCmd.size() * sizeof(wchar_t)),
                       &written, nullptr);
    if (ok) session->currentFile = FileToLoad;
    return ok ? 1 : 0;
}

int __stdcall ListLoadNext(HWND ParentWin, HWND PluginWin,
                           char* FileToLoad, int ShowFlags) {
    std::wstring wide = Utf8ToWide(FileToLoad);
    return ListLoadNextW(ParentWin, PluginWin, &wide[0], ShowFlags);
}

void __stdcall ListCloseWindow(HWND ListWin) {
    PreviewSession* session = nullptr;
    {
        std::lock_guard<std::mutex> lock(g_sessionsMutex);
        auto it = g_sessions.find(ListWin);
        if (it != g_sessions.end()) {
            session = it->second;
            g_sessions.erase(it);
        }
    }
    TerminateSession(session);
    DestroyWindow(ListWin);
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
    // TC použije tento string k rozhodnutí, zda plugin nabídnout
    const char* s = "EXT=\"DOCX\"|EXT=\"DOC\"|EXT=\"DOCM\"|"
                    "EXT=\"XLSX\"|EXT=\"XLS\"|EXT=\"XLSM\"|EXT=\"XLSB\"|"
                    "EXT=\"PPTX\"|EXT=\"PPT\"|EXT=\"PPTM\"|"
                    "EXT=\"RTF\"|EXT=\"VSDX\"|EXT=\"MSG\"";
    strncpy_s(DetectString, maxlen, s, _TRUNCATE);
}

int __stdcall ListSendCommand(HWND ListWin, int Command, int Parameter) {
    // Zatím nepodporujeme zoom/print/copy - rezerva pro budoucno
    return 0;
}

// ---------------------------------------------------------------------------
// DllMain
// ---------------------------------------------------------------------------

BOOL APIENTRY DllMain(HINSTANCE hInst, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        g_hInstance = hInst;
        DisableThreadLibraryCalls(hInst);
    } else if (reason == DLL_PROCESS_DETACH) {
        // Pokud TC zapomene zavřít okna, ukliď
        std::lock_guard<std::mutex> lock(g_sessionsMutex);
        for (auto& kv : g_sessions) TerminateSession(kv.second);
        g_sessions.clear();
    }
    return TRUE;
}
