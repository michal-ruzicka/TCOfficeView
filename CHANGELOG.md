# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-05-15

First working release.

### Added

- Total Commander Lister plugin (`tcoffice.wlx` / `tcoffice.wlx64`) that
  previews MS Office documents (DOCX, DOC, DOCM, XLSX, XLS, XLSM, XLSB,
  PPTX, PPT, PPTM, RTF, VSDX, MSG) inside TC's F3 / Quick View pane via
  the Windows Preview Handlers registered by Office.
- Out-of-process host (`tcoffice_host.exe` / `tcoffice_host_x86.exe`)
  that hosts the preview handler under COM STA; a crash in the host does
  not bring down Total Commander.
- Bitness-matched host launch: the 64-bit plugin spawns
  `tcoffice_host.exe`, the 32-bit plugin spawns `tcoffice_host_x86.exe`.
- Diagnostic logging to `%TEMP%\tcoffice_host.log`, controllable at
  compile time via the `HOST_LOG_ENABLED` macro in `tcoffice_host.cpp`.
- CMake-based build driven by `build.cmd` that produces both bitnesses
  and packages everything into `dist\tcoffice.v<version>.zip` — a Total
  Commander auto-install bundle with `pluginst.inf` and all repo-root
  Markdown docs. The version is read from `pluginst.inf` as the single
  source of truth.
- Static C/C++ runtime linkage so the artifacts have no `vcruntime*.dll`
  dependency.

<!--
Compare URLs go here once the project has a public remote, e.g.:

[Unreleased]: https://github.com/<owner>/tcoffice/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/<owner>/tcoffice/releases/tag/v0.1.0
-->
