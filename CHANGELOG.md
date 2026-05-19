# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Opt-in **full render mode** for Word documents, configured per
  application under a new `[Mode]` INI section. When `Mode/Word=full`
  is set, the plugin drives a Word.Application instance via OLE
  Automation and reparents its main window into the Lister pane,
  giving Print Layout rendering (proper pagination, headers and
  footers) instead of the simplified Web Layout produced by the
  Preview Handler. Cold start is slower (~2–4 s) and memory use is
  higher (~100–300 MB), so the default remains `quick`. On any
  failure (missing Office, protected document, COM activation
  refused, …) the plugin transparently falls back to quick mode.
- Per-document preview tweaks applied in full mode:
  - `ActiveWindow.View.Type = wdPrintView` — Print Layout view.
  - `ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit` — zoom is
    fitted to the page width and re-fits automatically when the
    Lister pane is resized (important for the narrow Quick View
    [Ctrl+Q] pane).
  - `ActiveWindow.DisplayRulers = False` — rulers hidden.
  - `Document.Protect(wdAllowOnlyReading)` — read-only enforcement
    on top of the `ReadOnly=True` open flag, so typing in the
    document is blocked and Ctrl+S has nothing to save.
- INI keys `[Mode] Excel` and `[Mode] PowerPoint` are reserved (also
  defaulting to `quick`); both currently behave as `quick` regardless
  of the configured value — full-mode handlers for them are planned.

### Notes

- Full mode deliberately changes **no Application-level Word setting**
  (status bar, ribbon state, default zoom, …). Office persists those
  to the user's profile on Quit, so changing them in our preview
  process would leak into the user's standalone Word. Only window-,
  view- and document-scoped settings are touched.
- Modern Microsoft 365 Word no longer exposes `Application.Hwnd`
  through IDispatch. When it returns `DISP_E_UNKNOWNNAME` the plugin
  falls back to `EnumWindows` looking for an `OpusApp` window whose
  title contains the document filename.

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
