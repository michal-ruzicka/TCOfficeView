# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-05-15

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

## [0.1.0] - 2026-05-15

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

<!--
Compare URLs go here once the project has a public remote, e.g.:

[Unreleased]: https://github.com/<owner>/TCOfficeView/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/<owner>/TCOfficeView/releases/tag/v0.1.0
-->
