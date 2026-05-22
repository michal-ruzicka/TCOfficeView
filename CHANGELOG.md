# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **Full render mode** for Word, Excel and PowerPoint via a new
  `[Mode]` INI section. Four values are accepted per application:
  `quick` (Preview Handler, no button), `quick-switchable` (Preview
  Handler with overlay button — **new default**), `full` (OLE
  Automation, no button), `full-switchable` (OLE Automation with
  overlay button). When a full-mode value is set, the plugin launches
  a real Word / Excel / PowerPoint instance, opens the file
  read-only, and embeds its window into the Lister pane, giving the
  full Print Layout / grid / slide rendering instead of the
  simplified Preview Handler pipeline. Cold start is ~2–4 s;
  memory ~100–300 MB per instance. On any failure the plugin
  transparently falls back to quick mode.
- **Mode-switch overlay button** in the top-right of every Word /
  Excel / PowerPoint preview (shown when the configured mode is
  `quick-switchable` or `full-switchable`). One click toggles the
  current preview between quick and full mode without changing the
  INI default; the switch is per-preview only — the next file (or
  re-opening the same one) starts at the configured default. Button
  size and font scale with the monitor's DPI. Always hidden for file
  types with no full-mode handler (`.msg`, `.vsdx`) regardless of
  the setting.
- Per-app preview niceties in full mode:
  - **Word** — Print Layout view, page-width zoom that auto-refits
    on Lister resize, rulers hidden, runtime read-only enforcement
    so typing in the preview is blocked.
  - **Excel** — zoom locked to 100% on load; initial frame size
    matches the Lister pane. Two intrinsic limitations of embedding
    Excel as a child window via OLE Automation, which no combination
    of style edits, window-message synthesis or activation tricks
    convinced Excel to overcome:
      1. Excel does not relayout when the pane is later resized;
         reopening the Lister at the desired size triggers a fresh
         correct layout.
      2. Interactive mouse input (cell selection, sheet-tab clicks,
         most ribbon buttons) is unreliable, because Excel gates much
         of that processing on being the foreground top-level window —
         and a reparented child of a foreign process never is.  Quick
         mode is unaffected; for interactive work, use quick mode.
         Full Excel mode is best treated as a read-only viewer.
  - **PowerPoint** — slide zoom auto-refits to the Lister pane on
    every resize. PowerPoint's main window appears briefly on
    screen before being embedded (a short visible flash) because
    PowerPoint cannot be told to run invisibly; Word and Excel
    embed silently.

### Changed

- Moved the `AppKind` enum declaration before `HostState` so that
  `HostState::currentFileApp` can be initialised at compile time without
  relying on a forward-declaration hack.

### Fixed

- **Preview handlers were activated in-process, which violates the
  `IPreviewHandler` contract.**  `LoadHandlerSta` used
  `CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER`, so COM happily loaded any
  handler that registered an `InProcServer32` (the Windows built-in MAPI
  Mail Previewer for `.msg`, in particular) directly into our process.
  Such handlers complete activation, initialisation and `SetWindow`, then
  fail `DoPreview` with `E_FAIL` because Microsoft's `IPreviewHandler`
  documentation states a handler must be hosted in its own background
  process.  Activation is now `CLSCTX_LOCAL_SERVER` only, matching the
  reference preview-host sample and Windows Explorer's preview pane —
  COM now routes through the AppID's `DllSurrogate` (typically
  `prevhost.exe`), and MSG previews finally render.
- **`CoInitializeSecurity` is now established at host start-up.**  Handlers
  that run in Windows' low-integrity `prevhost.exe` surrogate (notably the
  Outlook MSG preview handler) need an explicit process-wide COM security
  blanket to perform cross-process calls back to the host; without it
  `DoPreview` returns `E_FAIL`.  The new call matches the settings used by
  Windows Explorer's own preview pane (`RPC_C_AUTHN_LEVEL_DEFAULT`,
  `RPC_C_IMP_LEVEL_IMPERSONATE`).
- **Preview handlers requiring an `IPreviewHandlerFrame` site failed `DoPreview`.**
  Outlook's MSG preview handler (and other handlers hosted in `prevhost.exe`)
  ask the host for `IPreviewHandlerFrame` and `IOleWindow` via a site object
  installed through `IObjectWithSite::SetSite`; without it `DoPreview` returns
  `E_FAIL` and the preview never renders.  A new `PreviewHostSite` class
  implements `IPreviewHandlerFrame` (no-op accelerator translation),
  `IServiceProvider` (routing `IPreviewHandlerFrame` / `IOleWindow` lookups
  back to itself) and `IOleWindow` (returning the render window's HWND).
  The site is installed before any `IInitializeWith*::Initialize` call —
  Outlook's MSG handler queries the site from inside `Initialize` and aborts
  the load if it isn't ready yet — and cleared again on unload.  The call
  order in `LoadHandlerSta` now matches Microsoft's recommended sequence:
  CoCreateInstance → QI IPreviewHandler → SetSite → Initialize → SetWindow
  → DoPreview.
- **Preview handlers initialised only via `IInitializeWithItem` could not load.**
  `LoadHandlerSta` previously tried only `IInitializeWithFile` and
  `IInitializeWithStream`.  Outlook's MSG preview handler rejects both
  (`E_NOINTERFACE`) and supports exclusively the modern Shell-item-based
  `IInitializeWithItem`, which the loader now tries as a third fallback
  (`SHCreateItemFromParsingName` → `IInitializeWithItem::Initialize`).
- **Preview handlers missed when registered outside the direct extension key.**
  `FindPreviewHandlerClsid` previously looked for the preview handler CLSID
  only directly on the extension key and via the extension's default ProgID.
  This missed handlers registered in any of the other places Windows
  Explorer's preview pane consults — notably the MSG handler, which Outlook
  registers under a versioned ProgID in `HKCR\.msg\OpenWithProgids`
  (e.g. `Outlook.File.msg.16`) rather than as the default ProgID. The
  lookup now follows the full Shell association chain that Explorer uses:
  direct shellex key → default ProgID → each ProgID in `OpenWithProgids` →
  `HKCR\SystemFileAssociations\<ext>` → `HKCR\SystemFileAssociations\<PerceivedType>`.
  MSG and similar previews now work on systems where Windows Explorer already
  shows them.
- **Plugin deadlock on Lister close / TC exit.** `TerminateSession` used an
  infinite timeout when writing `CLOSE` to the host pipe. If the host's pipe
  reader thread had already died (e.g. access violation), the write would
  block forever, freezing Total Commander. The write now uses a 3-second
  timeout and cancels the pending I/O if it expires.
- **Race condition in plugin `ChildWndProc` (use-after-free).** The resize
  timer handler looked up the session under `g_sessionsMutex`, released the
  lock, and then wrote to `session->hPipe`. Between releasing the lock and
  the write, `ListCloseWindow` on another thread could erase and `delete`
  the session. The mutex is now held for the entire pipe write.
- **Race condition in host `HostLog` critical-section initialisation.**
  `InitializeCriticalSection` was called lazily on the first `HostLog`
  invocation without any synchronisation. Two threads (STA + pipe reader)
  could race and initialise the same `CRITICAL_SECTION` twice, corrupting
  it. The CS is now initialised once in `wmain` before the pipe reader
  thread is created.
- **Pipe buffer exhaustion after many file switches.** The host sends
  `OK\n` / `ERR …\n` replies for every `LOAD` and `CLOSE`, but the plugin
  never read them. After roughly 5 000 `ListLoadNextW` calls in a single
  Lister session the 64 KB pipe buffer filled up and subsequent writes
  blocked indefinitely. A background drain thread now continuously reads
  and discards host responses.
- **Handle leaks in the host process.** The log file handle, the logging
  `CRITICAL_SECTION`, and the Office `JobObject` handle were never closed
  or destroyed before process exit. They are now cleaned up in `wmain`
  after the message loop terminates.
- **Potential `CreateProcessW` command-line mutation.** `CreateProcessW`
  requires a writable command-line buffer. `cmdLine.data()` returns
  `const wchar_t*` in pre-C++17 standards, which is undefined behaviour.
  Changed to `&cmdLine[0]`, which is writable in all C++ versions.
- **Missing null check in `ListLoadNextW`.** Added an explicit guard against
  a `nullptr` `FileToLoad` argument.
- **Re-entrancy crash when switching modes rapidly (PowerPoint).** The retry
  loops inside `LoadWordFullSta`, `LoadExcelFullSta` and
  `LoadPowerPointFullSta` contained `PeekMessage` dispatch loops that pumped
  the STA message queue. While a load was still in progress, a second
  `WM_HOST_SWITCH_MODE` message could be dispatched recursively, causing
  re-entrant COM calls and eventual heap corruption or access violation. The
  `PeekMessage` loops have been replaced with plain `Sleep`, and a
  `loadingInProgress` guard now serialises `LoadFileWithModeSta` calls so
  that a second switch is deferred until the first one finishes.
- **Document not closed before quick-mode fallback detach.** When full-mode
  loading failed and the host fell back to quick mode, the Office document
  was left open in the background while the window was detached. The
  appropriate close routine (`CloseWordDocumentSta`,
  `CloseExcelWorkbookSta`, `ClosePptPresentationSta`) is now called before
  detaching the window in the fallback path.
- **Crash when clicking the embedded Office window's close button.**
  `EmbedOfficeWindowSta` strips Win32 caption styles, but Office apps
  (Word, Excel, PowerPoint) render their own close button inside the
  ribbon / title bar. Clicking it sent `WM_SYSCOMMAND(SC_CLOSE)` or
  `WM_CLOSE` to the embedded HWND, which caused the Office app to destroy
  its own window while it was still a child of our render pane — leading
  to a crash or unexpected TC window closure. `SetWindowSubclass` cannot
  be used because Office runs out-of-process. Instead, a visible light
  grey overlay bar (`hwndCloseGuard`) is created as a child of the render
  pane, stretched across the full width and covering the top ~40 px
  (HiDPI-scaled) where the Office ribbon / title bar lives. The guard is
  forced to `HWND_TOP` so it sits above the Office window, painted with
  `LTGRAY_BRUSH`, and returns `HTCLIENT` from `WM_NCHITTEST` so all mouse
  clicks in that area are swallowed. This also blocks the context menu
  that appears on a right-click in the title area. The guard is destroyed
  when the Office window is detached or the render pane is torn down.

### Notes

- Full mode deliberately changes no Office Application-wide settings
  (status bar, ribbon state, default zoom, …) because Office
  persists those into the user's profile on Quit. Only per-window /
  per-document settings are touched, so your standalone Word, Excel
  and PowerPoint will start up exactly as you left them.
- Every Office process spawned by the plugin is assigned to a Job
  Object with `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`. When the host
  process exits for any reason (clean shutdown, plugin-side
  `TerminateProcess` after the host failed to honour `CLOSE` within
  2 seconds, Windows shutdown), the kernel kills any still-alive
  Office processes with it — so Excel in particular can no longer
  be left orphaned and reported as a "hung application" at the next
  system restart.

## [1.0.0] – 2026-05-18

First public binary release.

### Added

- GitHub Actions CI workflow (`.github/workflows/build.yml`) that builds
  both bitnesses on every push and pull request using `build.cmd` and
  uploads the resulting ZIP as a downloadable workflow artifact. Releases
  are still published manually after local GPG signing.
- Build status badge in `README.md`.

### Changed

- Repository layout reorganised: all C++ sources, INI/INF/manifest files
  and `CMakeLists.txt` moved under `src/`. The repo root now holds only
  the Markdown documentation and `build.cmd`. The auto-install ZIP
  layout produced by `build.cmd` to `dist/` is unchanged.
- Contributor docs split out of `README.md` into a dedicated
  [`CONTRIBUTING.md`](CONTRIBUTING.md) at the repo root. The README now
  focuses on end-user installation, usage and configuration.
- `CONTRIBUTING.md` now includes a *Release Process* section describing
  the local-build + GPG-sign + manual GitHub upload workflow.

## [0.3.0] – 2026-05-18

### Added

- Fallback informational pane shown in place of a preview when no working
  preview handler is available — most commonly because MS Office is not
  installed, but also when a registered handler fails to load. The pane
  explains the situation and lists the file's basic metadata (name, path,
  extension, size, created / modified / last-accessed timestamps) so the
  Lister never just shows an empty panel.
- Font in the fallback panel is DPI-aware — it is configured in points and
  scaled to the actual monitor DPI when the panel is created.
- INI keys `[FallbackUI] FontFamily` and `[FallbackUI] FontSize` to
  override the font. When `FontFamily` is empty (the default) the plugin
  auto-picks the first installed font from a preference list: Aptos Mono
  → Consolas → Cascadia Mono → Lucida Console → Courier New.

### Changed

- Update documentation ([README.md](README.md), [CLAUDE.md](CLAUDE.md)).

### Fixed

- Non-ASCII characters in wide string literals (em-dashes, accented
  characters) no longer display as mojibake in the fallback panel on
  systems whose ANSI code page is not UTF-8. MSVC is now invoked with
  `/utf-8` so source files are interpreted as UTF-8 regardless of the
  system locale.

## [0.2.0] – 2026-05-15

### Added

- Runtime configuration via `TCOfficeView.ini`. The host looks for a per-user
  override under `%APPDATA%\GHISLER\TCOfficeView.ini` first and falls back to
  the system-wide INI shipped next to the host EXE; the first existing file
  wins (no per-key merging).
- INI-controlled diagnostic logging. `[Logging] LogPath=...` enables the
  log; an empty value (the default) disables it. Environment variables
  inside the path (`%TEMP%`, `%LocalAppData%`, …) are expanded, and any
  missing parent directories are created on first write.

### Changed

- Diagnostic logging is now **off by default**. The previous build wrote
  to `%TEMP%\TCOfficeViewHost.log` unconditionally whenever
  `HOST_LOG_ENABLED` was compiled in. That macro is still present as a
  compile-time kill-switch, but at run time logging is gated by the INI
  setting above.

## [0.1.0] – 2026-05-15

First working release.

### Added

- Total Commander Lister plugin (`TCOfficeView.wlx` / `TCOfficeView.wlx64`) that
  previews MS Office documents (DOCX, DOC, DOCM, XLSX, XLS, XLSM, XLSB,
  PPTX, PPT, PPTM, RTF, VSDX, MSG) inside TC's F3 / Quick View pane via
  the Windows Preview Handlers registered by Office.
- Out-of-process host (`TCOfficeViewHost.exe` / `TCOfficeViewHost_x86.exe`)
  that hosts the preview handler under COM STA; a crash in the host does
  not bring down Total Commander.
- Bitness-matched host launch: the 64-bit plugin spawns
  `TCOfficeViewHost.exe`, the 32-bit plugin spawns `TCOfficeViewHost_x86.exe`.
- Diagnostic logging to `%TEMP%\TCOfficeViewHost.log`, controllable at
  compile time via the `HOST_LOG_ENABLED` macro in `TCOfficeViewHost.cpp`.
- CMake-based build driven by `build.cmd` that produces both bitnesses
  and packages everything into `dist\TCOfficeView.v<version>.zip` — a Total
  Commander auto-install bundle with `pluginst.inf` and all repo-root
  Markdown docs. The version is read from `pluginst.inf` as the single
  source of truth.
- Static C/C++ runtime linkage so the artifacts have no `vcruntime*.dll`
  dependency.

[Unreleased]: https://github.com/michal-ruzicka/TCOfficeView/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/michal-ruzicka/TCOfficeView/compare/v0.3.0...v1.0.0
[0.3.0]: https://github.com/michal-ruzicka/TCOfficeView/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/michal-ruzicka/TCOfficeView/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/michal-ruzicka/TCOfficeView/releases/tag/v0.1.0
