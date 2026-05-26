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
- `[Mode] Word=` / `Excel=` / `PowerPoint=` — four values accepted:
  `quick-switchable` (default), `quick`, `full-switchable`, `full`.
  `BaseMode()` extracts the actual loader mode (Quick/Full);
  `IsSwitchable()` says whether to show the overlay button.
- `[AutoFallback] Word=` / `Excel=` / `PowerPoint=` — per-app toggle
  (default `true`) for the quick→full auto-fallback path described
  below. Accepts true/false, yes/no, on/off, 1/0. Read into
  `g_autoFallbackWord` / `g_autoFallbackExcel` /
  `g_autoFallbackPowerPoint` and read back at run time via
  `IsAutoFallbackEnabled(app)`.

The fallback panel is a child `EDIT` control inside `hwndRender`,
created on any failure from `FindPreviewHandlerClsid` /
`CoCreateInstance` / `Initialize*` / `QI IPreviewHandler` /
`SetWindow` / `DoPreview`. `LoadHandlerSta` still returns `S_OK` to the
plugin in that case — the file was "shown", just via the fallback. The
panel is sized in `ResizeHandlerSta` and torn down in
`UnloadHandlerSta`.

`LoadHandlerSta` is a thin wrapper around `TryLoadHandlerSta`. The
`Try` variant returns the actual failure `HRESULT` instead of
silently showing the fallback panel; the wrapper calls
`ShowFallbackSta(path, hr)` on failure. The split lets the caller
(`LoadFileWithModeSta`'s quick path) detect quick-mode failure
*before* the user sees anything and silently retry in full mode for
Office files — see "Auto-fallback quick→full" below.

### Auto-fallback quick→full for SharePoint multi-tenant files

Files synced from SharePoint sites in non-primary Microsoft 365
tenants typically fail to render in quick mode (preview handler
returns `E_FAIL` from `DoPreview`, or `E_NOTIMPL` from
`IInitializeWithFile`). This is an Office-side limit, not ours: the
preview handlers run in the `prevhost.exe` surrogate process and
cannot authenticate cross-tenant. The same files refuse to preview
in **Explorer's Alt+P / Preview Pane** for the same reason —
verify there before suspecting our code if the issue resurfaces.

The real Office app (`winword.exe` / `excel.exe` / `powerpnt.exe`)
runs as the user with full auth tokens and usually opens the same
file fine. The auto-fallback path in `LoadFileWithModeSta` exploits
that: when `TryLoadHandlerSta` returns failure for an Office file
the code attempts `LoadXxxFullSta` immediately and only shows the
fallback panel if that also fails.

Gated on five conditions, all of which must hold:

1. `app` is Word, Excel or PowerPoint.
2. `SelectMode(app) == Mode::QuickSwitchable` — the user has opted
   into per-file mode switching for that app, so the implicit
   switch matches expectation. `Mode::Quick` is treated as "quick
   only, no fallback" by design.
3. `IsAutoFallbackEnabled(app)` is true (the `[AutoFallback]` INI
   toggle).
4. `!fellThroughFull` — we did not arrive at the quick path via a
   full-mode failure ourselves. Without this guard a file that
   fails in both modes would loop forever.
5. `ReadFileZoneInfo(origPath).zoneId < 3` — file is NOT MOTW-
   blocked. MOTW is by far the most common reason quick fails on
   an otherwise-healthy handler, and the typical fix is a single
   Unblock click that lets the file render in quick mode — which
   is simpler and lighter than full mode. We prefer to surface the
   Unblock button and let the user decide rather than silently
   start a heavyweight Office process. The cross-tenant case (file
   is MOTW AND would also fail quick after unblock) is handled
   naturally by the second LOAD that the Unblock click triggers:
   the file is no longer MOTW so this gate doesn't fire, and the
   SharePoint cross-tenant failure proceeds cleanly into auto-
   fallback.

The explicit-full path (`mode == Mode::Full` above) deliberately
does NOT have a corresponding MOTW pre-check. Office's Automation
pipeline applies Protected View on its own for MOTW files (yellow
"Enable Editing" banner, content visible in a read-only sandbox),
which is the right outcome when the user has asked for full mode.

On auto-fallback success the code sets
`g_state.currentLoadedMode = Mode::Full` so the mode-switch button
shows `→ Quick` (correctly reflecting the current state) and the
button click handler swaps the right direction.

`LoadFileWithModeSta` takes an `allowAutoFallback` parameter
(default `true`). The initial-LOAD path (`LoadFileSta`, the
`WM_HOST_LOAD` pipe command) leaves it at the default so SharePoint
multi-tenant documents "just work". The user-initiated
`WM_HOST_SWITCH_MODE` handler passes `false`: when the user clicks
`→ Quick` on a document that auto-fell-back to full at load time,
the click is treated as an explicit "I want quick mode here"
instruction. Without this guard the click would just trigger
auto-fallback again and feel like the button did nothing. The
fallback panel is shown instead and `UpdateModeButtonSta` flips the
overlay button to `→ Full` so the user can return deliberately.

`WM_HOST_UNBLOCK_AND_RELOAD` keeps the default `true` — after the
ADS strip the file is no longer MOTW and a SharePoint cross-tenant
failure on the reload should still benefit from auto-fallback.

The Unblock button itself is added by `ShowFallbackSta` whenever
the fallback panel is shown for a MOTW-blocked file
(`ReadFileZoneInfo(path).zoneId >= 3`). That naturally restricts
its visibility to the quick-mode failure path — the only way the
fallback panel is displayed now that the full-mode pre-check is
gone. If the user is configured for full mode (or auto-fallback
landed them there) Office shows its own Protected View UI and the
Unblock button never appears.

The generic-failure branch of `BuildFallbackText` (the non-MOTW,
non-`REGDB_E_CLASSNOTREG` path) calls `ClassifyByExtension(path)`
and, for Office files, expands the message with the SharePoint
multi-tenant cause and an Explorer-Alt+P-parity note. When the
file's app is in a `IsSwitchable(...)` mode the message also
includes a "click → Full" hint pointing at the mode-switch button.
The hint is suppressed when the user has configured plain `quick`
(non-switchable) since the button would not be visible.

## Mode dispatch (quick vs full)

`Mode` has four values: `Quick`, `QuickSwitchable`, `Full`,
`FullSwitchable`. `BaseMode(m)` returns `Quick` or `Full`; `IsSwitchable(m)`
returns whether the mode-switch button should be shown.

`LoadFileSta` is the top-level LOAD entry point. It classifies the file by
extension into `AppKind { Other, Word, Excel, PowerPoint }`, reads the
config via `SelectMode(app)` (which may return a Switchable value), and
calls `LoadFileWithModeSta(path, app, BaseMode(cfg))`. The worker only
ever receives `Quick` or `Full`.

`LoadFileWithModeSta` dispatches:

- `Mode::Full` + `AppKind::Word` → `LoadWordFullSta`
- `Mode::Full` + `AppKind::Excel` → `LoadExcelFullSta`
- `Mode::Full` + `AppKind::PowerPoint` → `LoadPowerPointFullSta`
- Anything else → `LoadHandlerSta` (the original preview-handler path).

On any full-mode failure the dispatcher silently falls back to quick mode.

`LoadFileWithModeSta` stores `currentFile`, `currentFileApp` and
`currentLoadedMode` (`Quick` or `Full`) into `HostState`. These are used by:

- the mode-switch overlay button to know what to re-load in the opposite mode;
- `UpdateModeButtonSta` to pick the label (`→ Full` / `→ Quick`) and decide
  visibility: `IsSwitchable(SelectMode(currentFileApp))` — hidden when the
  configured mode is non-switchable or the file type is `Other`.

The button's click handler in `RenderWndProc` posts `WM_HOST_SWITCH_MODE`
to the STA window; the STA handler calls
`LoadFileWithModeSta(currentFile, currentFileApp, opposite-mode)`. The
switch is per-preview only — the next LOAD command from the plugin DLL
goes through `LoadFileSta`, which picks up the INI-configured default again.

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

### Office process lifetime

Office processes (Word / Excel / PowerPoint) spawned via COM
(`CLSCTX_LOCAL_SERVER`) normally exit on their own once the last
client COM reference is released, but that signal can get lost when
the host dies abruptly (plugin-side `TerminateProcess` after a
hanging Quit, Windows shutdown, host crash). To make the lifetime
deterministic, the host creates a Job Object with
`JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE` in `wmain` and assigns each
Office process to it right after `EmbedOfficeWindowSta` (via the
HWND → `GetWindowThreadProcessId` → `OpenProcess` →
`AssignProcessToJobObject` sequence). When the host process exits
through *any* path, the kernel closes the Job handle and the
`KILL_ON_JOB_CLOSE` flag terminates every assigned Office process.

The Job handle is intentionally **not** closed in `wmain`'s cleanup
code: closing it triggers the kill immediately, which would race
the `Application.Quit` calls in `UnloadXxxFullSta`. Letting the OS
close it on process exit means the clean-quit path gets a chance
to finish first and the Job is only the safety net for unclean
exits.

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
