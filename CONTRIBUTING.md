# For Contributors

**This document is for developers who want to build, modify or extend the plugin.**

The project lives at <https://github.com/michal-ruzicka/TCOfficeView>. Bug
reports and patches are welcome via the project's
[issue tracker](https://github.com/michal-ruzicka/TCOfficeView/issues) and
[pull requests](https://github.com/michal-ruzicka/TCOfficeView/pulls).

## Prerequisites

- [**Visual Studio 2026 Build Tools**](https://visualstudio.microsoft.com/) with
  the *Desktop development with C++* workload. The full Visual Studio IDE works
  too, but is not required.
- [**CMake 3.20+**](https://cmake.org/). Often bundled with the Build Tools
  workload.

No .NET SDK, no extra runtime — the project is pure C++ / Win32 / COM,
linked against the static MSVC runtime.

## Build

From a Developer Command Prompt (or any shell where `cmake` and MSVC are
on `PATH`):

```cmd
build.cmd
```

The script drives CMake for both bitnesses (x86 and x64) using the sources
under `src\` and packages the artifacts into
`dist\TCOfficeView.v<version>.zip` — the same auto-install bundle
described in *Installation*. The version is read from `src\pluginst.inf`
as the single source of truth; bump it there before tagging a release.

## Repo Layout

- `src\` — C++ sources, INI/INF/manifest, `CMakeLists.txt`
- `build.cmd` — top-level build driver (the only file outside `src\`
  that the build touches)
- `build\` — CMake out-of-source build trees (gitignored)
- `dist\` — released ZIP bundles (gitignored)
- Markdown files (README, CONTRIBUTING, CHANGELOG, LICENSE, TODO,
  CLAUDE) live at the repo root and are bundled into every release ZIP.

## Architecture

Two artifacts, one process boundary between them:

```
Total Commander process
└── TCOfficeView.wlx[64]   (plugin DLL, C++/Win32)
    ├── creates a child HWND inside TC's Lister pane
    └── spawns TCOfficeViewHost.exe per Lister session, talks to it
              │   over a named pipe
              ▼
TCOfficeViewHost.exe       (host EXE, C++/Win32/COM)
├── CoInitializeEx(COINIT_APARTMENTTHREADED)
├── CoCreateInstance(CLSID for the file's extension)
├── IInitializeWithFile / IInitializeWithStream::Initialize(path)
├── IPreviewHandler::SetWindow(child HWND) + DoPreview()
└── Win32 message loop on the STA thread
```

The plugin runs the host in a separate process for three reasons:

1. **Stability.** Office preview handlers occasionally crash. A crash in
   the host does not bring down Total Commander. (Windows Explorer uses
   the same isolation pattern via `prevhost.exe`.)
2. **STA threading.** `IPreviewHandler` requires a single-threaded
   apartment. Guaranteeing STA inside a DLL loaded by an arbitrary host
   is fragile; an own EXE makes the apartment explicit.
3. **Bitness flexibility.** Preview handlers are registered per bitness.
   `CoCreateInstance` with `CLSCTX_LOCAL_SERVER` lets Windows bring up
   the handler in the matching bitness regardless of which bitness of
   plugin and host are loaded.

Deeper design notes — the rationale for ditching the original .NET host,
the threading invariants, the `WM_SIZE` coalescing, the cross-process
`SetParent` trick that broke deadlocks against TC's UI thread, the
`PostMessage`-only worker → STA path that sidesteps
`RPC_E_CANTCALLOUT_ININPUTSYNCCALL` — live in [CLAUDE.md](CLAUDE.md).

## Plugin ↔ Host Wire Protocol

Text-based, UTF-16 LE, lines terminated by `\n`, transported over a
message-mode named pipe.

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

`RESIZE` is fire-and-forget — the host does not acknowledge it, because
the plugin coalesces `WM_SIZE` events through a 50 ms timer and a steady
stream of resize messages would otherwise saturate the pipe during a
drag.

## Known Limitations

- **First-preview latency** – roughly 200–800 ms (cold start of the
  Office COM server). Subsequent previews are 100–300 ms.
- **Memory** — each host process holds ~30–80 MB while a preview is
  active. Successive previews in the same Lister session reuse the same
  host; cross-session pooling is not implemented.
- **`ListSendCommand` is a stub.** Print / find / copy from TC's Lister
  menu are not yet wired through to the host. Most preview handlers
  expose their own context menus inside the preview area, so this is
  rarely noticed in practice.

## License

This project is licensed under the Apache License 2.0 — see
[LICENSE.md](LICENSE.md) for the full text. Source files carry SPDX
identifiers (`SPDX-License-Identifier: Apache-2.0`). Release notes are in
[CHANGELOG.md](CHANGELOG.md).
