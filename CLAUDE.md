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
  `IPreviewHandler::*`, `IDispatch::Invoke` on Word / Excel /
  PowerPoint Application) must
  happen on the STA thread. The pipe reader runs on a worker thread and
  dispatches via `PostMessage` to a message-only window on the STA
  thread. `SendMessage` would put the STA into the input-synchronous
  state, where COM refuses outgoing calls with
  `RPC_E_CANTCALLOUT_ININPUTSYNCCALL` (0x8001010D) — the bug that broke
  the first working build.
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
- `[Mode] Word=` / `Excel=` / `PowerPoint=` — `quick` (default) selects
  the Preview Handler path; `full` selects the OLE Automation path. All
  three apps are implemented in full mode.

The fallback panel is a child `EDIT` control inside `hwndRender`,
created on any failure from `FindPreviewHandlerClsid` /
`CoCreateInstance` / `Initialize*` / `QI IPreviewHandler` /
`SetWindow` / `DoPreview`. `LoadHandlerSta` still returns `S_OK` to the
plugin in that case — the file was "shown", just via the fallback. The
panel is sized in `ResizeHandlerSta` and torn down in
`UnloadHandlerSta`.

## Mode dispatch (quick vs full)

`LoadFileSta` is the top-level LOAD entry point. It classifies the file by
extension into `AppKind { Other, Word, Excel, PowerPoint }`, looks up the
matching `Mode` from `[Mode]`, and dispatches:

- `Mode::Full` + `AppKind::Word` → `LoadWordFullSta`
- `Mode::Full` + `AppKind::Excel` → `LoadExcelFullSta`
- `Mode::Full` + `AppKind::PowerPoint` → `LoadPowerPointFullSta`
- Anything else → `LoadHandlerSta` (the original preview-handler path).

On any full-mode failure (Office missing, COM activation refused,
document corrupted, …) the dispatcher silently falls back to the quick
path so the user always gets *some* preview.

Full-mode uses **OLE Automation via raw `IDispatch`** — no type library
`#import`, no MFC. The thin helpers `DispGetId`, `DispCall`,
`DispGetProperty`, `DispPutBool`, `DispPutI4`, `DispGetDispatchProperty`,
`DispGetHwndProperty` wrap the `GetIDsOfNames` + `Invoke` dance, and are
shared across all three apps.

**Only one Office app can be embedded at a time** — the render pane is a
single embed point. Each `LoadXxxFullSta` calls `UnloadYyyFullSta(true)`
on the other two apps before installing its own window. Within a Lister
session, switches between files of the *same* app reuse the running
Application instance (just close/open the document); switches across
apps pay the cold-start cost again.

Per-app COM details:

| App        | ProgID                  | Collection      | Open call signature                                   | Frame class    |
|------------|-------------------------|-----------------|-------------------------------------------------------|----------------|
| Word       | `Word.Application`      | `Documents`     | `Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)` | `OpusApp`      |
| Excel      | `Excel.Application`     | `Workbooks`     | `Open(FileName, UpdateLinks, ReadOnly)`               | `XLMAIN`       |
| PowerPoint | `PowerPoint.Application`| `Presentations` | `Open(FileName, ReadOnly, Untitled, WithWindow)`      | `PPTFrameClass`|

`Application.Visible` can be `False` for Word and Excel (the window
materialises only after we set it back to `True` for the reparent step).
**PowerPoint rejects `Visible = False`**, so its main window is always on
screen between `Presentations.Open` and our `SetParent` — accept the
brief flash.

`Document.Close` arguments differ:
- Word: `Close(SaveChanges = wdDoNotSaveChanges = 0)` (Long enum)
- Excel: `Close(SaveChanges = False)` (Variant Boolean)
- PowerPoint: `Close()` no args, with `Presentation.Saved = True` set
  beforehand to suppress the save prompt.

On CLOSE (Lister shutdown) all three Unload functions run; each is a
no-op if its slot is empty.

### Embedding the app window

For Word and Excel the main HWND does not materialise until
`Application.Visible = True`. For PowerPoint the HWND exists from
`CoCreateInstance` since `Visible = False` is not allowed. The embedding
sequence (parameterised on the app) is:

1. `CoCreateInstance(*.Application, IID_IDispatch)` — process is up.
2. Suppress dialogs: Word/PowerPoint use a `DisplayAlerts` enum (Word
   `wdAlertsNone = 0`, PowerPoint `ppAlertsNone = 1`); Excel uses a
   Boolean `DisplayAlerts = False`.
3. For Word and Excel only: `Visible = False`.
4. `Documents.Open` / `Workbooks.Open` / `Presentations.Open` with the
   per-app args listed above. `IDispatch::Invoke` takes positional args
   in **reverse order**; we don't try to skip ahead to optional
   parameters past the first few (that would require named-arg
   invocation).
5. For Word and Excel: set `Visible = True` so the main window
   materialises. (PowerPoint is already visible.)
6. Locate the HWND. **`Application.Hwnd` is gone in modern Microsoft
   365** (`GetIDsOfNames` returns `DISP_E_UNKNOWNNAME`). We try the
   property once for older Office, then fall back to `EnumWindows`
   looking for the app's frame class (`OpusApp` / `XLMAIN` /
   `PPTFrameClass`), filtered by a window-title substring (the document
   filename) to disambiguate from any pre-existing instance the user
   may have running.
7. `EmbedOfficeWindowSta`: `SetParent` into `hwndRender`, strip
   `WS_CAPTION|WS_THICKFRAME|WS_SYSMENU|WS_BORDER|WS_POPUP|...`, add
   `WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN`, `SetWindowPos` with
   `SWP_FRAMECHANGED`.
8. Store the HWND into the per-app slot (`hwndWordApp` /
   `hwndExcelApp` / `hwndPptApp`).

### Per-document preview tweaks (Word only)

After embedding, `ConfigureWordForPreviewSta` applies:

- `ActiveWindow.View.Type = wdPrintView` (= 3) — Print Layout.
- `ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit` (= 2) — page-width
  zoom that re-fits on window resize.
- `ActiveWindow.DisplayRulers = False`.
- `Document.Protect(wdAllowOnlyReading)` — runtime read-only on top of
  the `ReadOnly=True` open flag.

Excel: `ConfigureExcelForPreviewSta` sets `ActiveWindow.Zoom = 100`
(a Long percentage). We previously tried `Zoom = True` (Excel's "Fit
Selection" mode) but the resulting zoom depended on whatever range
Excel chose to fit and was visually unpredictable. Excel has no
auto-refit-on-resize equivalent of Word's `wdPageFitBestFit`.

Excel embedding has a **known limitation**: once Excel's frame is
embedded, the inner widgets (ribbon, `XLDESK`, sheet tabs, status
bar) lay out **once** at the initial frame size and do not relayout
on subsequent programmatic resizes. To make the initial layout match
the Lister pane we set `Application.Width`/`Height` (in points) just
before `Workbooks.Open` — Excel's frame size at workbook-open time is
what its layout engine commits to. Lister resizes after that point
adjust the Win32 frame (`SetWindowPos` in `ResizeOfficeFullSta`) but
Excel keeps drawing into the original area. Attempted workarounds
that did **not** work in our setup:

- `Application.Width`/`Height` toggled on resize — no effect on
  child widgets.
- `Application.WindowState = xlMaximized` after embed — blanks the
  ribbon area.
- `ActiveWindow.WindowState = xlMaximized` (maximize workbook
  inside the MDI client) — also blanks the ribbon and doesn't make
  the workbook follow the frame.
- Synthetic `WM_EXITSIZEMOVE` after `SetWindowPos` — Excel ignores it.
- `SWP_FRAMECHANGED` / `SWP_NOSENDCHANGING` flags on `SetWindowPos`
  — no effect.

The user-facing workaround is to close and reopen the Lister at the
desired size; that triggers a fresh `LoadExcelFullSta` which sets
`Application.Width`/`Height` before the workbook opens.

PowerPoint: `ConfigurePowerPointForPreviewSta` sets
`ActiveWindow.View.ZoomToFit = msoTrue`. Unlike Excel, this *is*
auto-refit: PowerPoint rescales the slide as the window changes size.

**Strict rule: we touch no `Application`-level Word setting.** Office
persists Application properties (status bar, ribbon state, default
zoom, recent files, …) into the user's profile on Quit, so changing
them here would silently rewrite the user's standalone-Word
preferences. Only window-, view- and document-scoped properties are in
scope.

## Future work (not blocking)

- Prefer `IInitializeWithStream` over `IInitializeWithFile` where supported;
  it would later let us preview files from inside archives.
- `ListSendCommand` is a stub. Print / find / copy from TC's Lister menu
  could be wired through to the host.
- Persistent host process pooling — currently one host per Lister session.
  Reuse via `ListLoadNextW` is implemented; cross-session pooling is not.
- Release workflow automation (GitHub Actions triggered by `v*` tags).
