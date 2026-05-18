# TCOfficeView — project notes for Claude

## What this is

A Total Commander **Lister plugin** (`.wlx` / `.wlx64`) that previews MS Office
documents (DOCX, XLSX, PPTX, RTF, VSDX, MSG, ...) inside Total Commander's
F3 / Quick View pane. The plugin does not parse the documents itself — it hosts
the Windows **Preview Handlers** that Office registers under
`HKCR\.<ext>\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}`. That's the same
COM contract Windows Explorer uses for its preview pane.

## Architecture

Two artifacts, one process boundary between them:

```
Total Commander process
└── TCOfficeView.wlx[64]                    (this repo: plugin DLL, C++/Win32)
    ├── exports ListLoad / ListLoadW / ListCloseWindow / ...
    ├── creates child HWND inside TC's Lister pane
    └── spawns TCOfficeViewHost.exe per session, talks to it over a named pipe
              │
              │  named pipe   (UTF-16 text protocol: LOAD / RESIZE / CLOSE)
              ▼
TCOfficeViewHost.exe                          (this repo: host EXE, C++/Win32/COM)
├── CoInitializeEx(COINIT_APARTMENTTHREADED)
├── CoCreateInstance(CLSID for the file's extension)
├── IInitializeWithFile::Initialize / IInitializeWithStream::Initialize
├── IPreviewHandler::SetWindow(child HWND) + DoPreview()
└── runs a Win32 message loop on the STA thread (required for COM)
```

### Why a separate host process?

1. **Stability** — Office preview handlers occasionally crash. A crash in the
   host process does not bring down Total Commander. (Windows Explorer uses
   the same isolation pattern via `prevhost.exe`.)
2. **STA threading** — `IPreviewHandler` requires a single-threaded apartment.
   Guaranteeing STA inside a DLL loaded by TC is fragile; an own EXE makes
   the apartment explicit and uncontested.
3. **Bitness matching** — Preview handlers are registered per bitness. A
   64-bit Total Commander with 32-bit Office (or vice versa) needs the host
   in the bitness that matches the *handler*, not TC. Shipping two host EXEs
   solves this.

### Why pure C++ / Win32 (no .NET)

The first sketch in this repo's initial commit used a .NET 8 host. We
deliberately removed that because:

- TC plugins are expected to be drop-in (unzip into a folder, point TC at it).
  A .NET runtime dependency breaks that.
- `.NET` startup adds ~150–300 ms to cold preview latency; native start is
  under 50 ms.
- `IPreviewHandler` is a C-first COM API. Hand-rolled `[ComImport]` interop in
  C# is more boilerplate than the equivalent C++ interface declarations.

## Repo layout

Markdown documentation and the build driver live at the repo root; all
build inputs (C++ sources, INI/INF/manifest, CMakeLists) live under
`src/`. `build.cmd` is the only file that crosses the boundary — it
reads from `src/` and writes to `build/` and `dist/` (both ignored).

```
README.md             - end-user docs (auto-installed into the ZIP)
CONTRIBUTING.md       - developer docs (auto-installed into the ZIP)
CHANGELOG.md          - Keep-a-Changelog formatted release notes
LICENSE.md            - Apache License 2.0 + copyright notice
TODO.md               - working list of upcoming work
CLAUDE.md             - this file
build.cmd             - convenience driver: x86 + x64, stages everything,
                        and packages it into dist/TCOfficeView.v<version>.zip
src/
  CMakeLists.txt      - builds TCOfficeView.wlx / .wlx64 and TCOfficeViewHost.exe
  TCOfficeView.cpp    - plugin DLL (TC Lister plugin entry points)
  TCOfficeView.def    - DLL exports
  listplug.h          - Lister plugin API (subset used here)
  TCOfficeViewHost.cpp - host EXE (COM preview handler host + fallback UI)
  app.manifest        - host EXE manifest (PerMonitorV2 DPI)
  pluginst.inf        - TC auto-installer metadata; canonical version source
  TCOfficeView.ini    - sample / system-wide config (logging, fallback font)
build/                - CMake out-of-source build trees (gitignored)
dist/                 - released TCOfficeView.v<version>.zip files (gitignored)
```

## Pipe protocol

UTF-16 LE, lines terminated by `\n`. The DLL is the pipe server, the host
is the client.

**Plugin → Host:**
```
LOAD <absolute-path>
RESIZE <width> <height>
CLOSE
```

**Host → Plugin:**
```
OK
ERR <message>
```

`RESIZE` is fire-and-forget — the host does not respond (resize fires at very
high frequency during window drag). The DLL coalesces `WM_SIZE` events via a
short timer to avoid stalling TC's UI on the pipe write.

## Things to keep in mind when editing

- **Both bitnesses must build.** 32-bit TC users exist. Don't add APIs that
  are 64-bit-only without guarding them.
- **Static CRT.** The plugin DLL links the static runtime
  (`MSVC_RUNTIME_LIBRARY = MultiThreaded`) so it has no `vcruntime*.dll`
  dependency. The host EXE links statically too, for the same reason.
- **No exceptions across COM boundaries.** Inside the host, COM HRESULTs
  are checked explicitly and turned into `ERR` responses; do not let C++
  exceptions escape into COM.
- **STA discipline in the host.** All COM calls (`CoCreateInstance`,
  `IPreviewHandler::*`) must happen on the STA thread. The pipe reader
  runs on a worker thread and dispatches via `PostMessage` to a
  message-only window on the STA thread. `SendMessage` would put the
  STA into the input-synchronous state, where COM refuses outgoing calls
  with `RPC_E_CANTCALLOUT_ININPUTSYNCCALL` (0x8001010D) — the bug that
  broke the first working build.
- **Don't block in `WM_SIZE`** in the plugin. The DLL runs inside TC's UI
  thread; a blocking pipe write would stall the resize loop.
- **Render in our process, not TC's.** The host creates its own child
  HWND and `SetParent`s it into the plugin's child window inside TC.
  Giving the preview handler our window (rather than TC's) keeps the
  handler's internal `SetParent` calls local to the host process and
  avoids cross-process win32k/COM deadlocks against TC's UI thread.
- **Source files are UTF-8.** CMakeLists passes `/utf-8` to MSVC so
  literals containing em-dashes / accented characters don't depend on
  the system ANSI code page.

## Build prerequisites

- Visual Studio 2026 Build Tools, workload: *Desktop development with C++*.
- CMake 3.20+.
- No .NET SDK. No other toolchain.

## Configuration & fallback UI

Runtime settings live in `TCOfficeView.ini`. The host looks for it under
`%APPDATA%\GHISLER\TCOfficeView.ini` first (per-user override) and falls
back to `<plugin install dir>\TCOfficeView.ini`. The first file that
exists wins entirely — no per-key merging. All values support
environment-variable expansion (`%TEMP%`, `%LocalAppData%`, …).

Sections currently honoured:

- `[Logging] LogPath=` — empty disables logging; non-empty path enables
  it. Missing parent directories are created on first write.
- `[FallbackUI] FontFamily=` / `FontSize=` — font for the read-only
  panel shown when no preview handler is usable for the file. Empty
  family auto-picks Aptos Mono → Consolas → Cascadia Mono → Lucida
  Console → Courier New. Font is DPI-aware (`GetDpiForWindow` +
  `MulDiv`).

The fallback panel is a child `EDIT` control inside `hwndRender`,
created on any failure from `FindPreviewHandlerClsid` /
`CoCreateInstance` / `Initialize*` / `QI IPreviewHandler` /
`SetWindow` / `DoPreview`. `LoadHandlerSta` still returns `S_OK` to the
plugin in that case — the file was "shown", just via the fallback. The
panel is sized in `ResizeHandlerSta` and torn down in
`UnloadHandlerSta`.

## Future work (not blocking)

- Prefer `IInitializeWithStream` over `IInitializeWithFile` where supported;
  it would later let us preview files from inside archives.
- `ListSendCommand` is a stub. Print / find / copy from TC's Lister menu
  could be wired through to the host.
- Persistent host process pooling — currently one host per Lister session.
  Reuse via `ListLoadNextW` is implemented; cross-session pooling is not.
- Full embedded mode instead of OLE preview (Word paged layout instead
  of web view), per-file-type configurable. See `TODO.md`.
- Project-level CI / GitHub release pipeline. See `TODO.md`.
