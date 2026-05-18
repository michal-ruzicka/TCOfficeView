# TCOfficeView: a Total Commander Lister Plugin that Previews Microsoft Office Documents

**A [Total Commander](https://www.ghisler.com/) Lister plugin that previews
Microsoft Office documents — Word, Excel, PowerPoint and more — directly in TC's
F3 / Quick View pane.** It hosts the Windows Preview Handlers that an MS Office
installation already registers, so the previews look exactly like the ones in
Windows Explorer or Outlook, without any document parsing of its own and without
a managed runtime.

**This plugin is not a replacement for MS Office software; a working MS Office
installation is required on the target computer** (see section [When MS Office
is not installed](#when-ms-office-is-not-installed) for more information). The
plugin then integrates document viewing (handled directly by the MS Office
applications) into Total Commander for quick and consistent work with MS Office
files alongside other file types such as images. If MS Office is not installed,
the plugin shows a small information panel with the file's basic metadata
instead of an empty preview, so it stays useful in any environment.

---

## Installation

The recommended way is the **Total Commander auto-installer:**

1. Download `TCOfficeView.vX.Y.Z.zip` from the latest release.
2. In Total Commander, navigate to the ZIP file and press **Enter**.
3. TC detects the bundled `pluginst.inf` and asks whether to install the
   plugin — confirm.
4. Press **F3** on a `.docx` / `.xlsx` / `.pptx` file to verify.

If you prefer to install manually, unzip the archive into a persistent
folder (for example `C:\Tools\TCOfficeView\`) and add `TCOfficeView.wlx`
(32-bit TC) or `TCOfficeView.wlx64` (64-bit TC) under **Configuration →
Options → Plugins → Lister plugins → Configure → Add**.

## Usage

There is no UI to configure at runtime — just press **F3** (or use Quick View
[**Ctrl+Q**]) on any supported file. The preview tracks the Lister pane size and
selecting a different file in the panel loads it into the same Lister session.

## Supported Formats

Format support depends on which Windows Preview Handlers are registered
on the machine. With Microsoft Office installed you typically get:

| Application | Extensions |
|-------------|------------|
| Word        | DOC, DOCX, DOCM, RTF |
| Excel       | XLS, XLSX, XLSM, XLSB |
| PowerPoint  | PPT, PPTX, PPTM |
| Visio       | VSD, VSDX |
| Outlook     | MSG |

The plugin advertises support for the union of these extensions. If a
specific handler is missing, the plugin shows the fallback information
panel for that file instead of failing silently.

> **Word documents render in Web Layout view.** All Office Preview
> Handlers (Word, Excel, PowerPoint) use a simplified embedded rendering
> pipeline optimised for speed and stability inside a host window, not
> the full editing UI. For Word this means a flowing Web Layout without
> page breaks, headers or footers; Excel shows a simplified grid;
> PowerPoint shows static slides without transitions. The plugin invokes
> the handler the same way Windows Explorer (Alt+R) does — the rendering 
> mode is hard-wired by Microsoft's handler implementation and there is no 
> API on `IPreviewHandler` to change it. A future “full embedded” mode that
> would drive a hidden Word/Excel/PowerPoint instance via OLE Automation
> instead of the Preview Handler is a possible future development.

> **MSG and VSDX caveat.** Recent Office / Outlook installs sometimes do
> not register the shell Preview Handler for `.msg` or `.vsdx` — New
> Outlook in particular drops the classic MAPI previewer. In those cases
> Windows Explorer's own preview pane is empty too, and the plugin falls
> back to the information panel.

## Configuration

Settings live in `TCOfficeView.ini`. The plugin reads it from two
locations and uses the **first one that exists** — values are not merged
across files:

1. `%APPDATA%\GHISLER\TCOfficeView.ini` — per-user override
2. `<plugin install dir>\TCOfficeView.ini` — system-wide default, shipped
   with the plugin

The shipped INI has every option commented out, so all defaults apply.
To customise, copy that file to `%APPDATA%\GHISLER\TCOfficeView.ini` and
edit the copy. Environment variables (`%TEMP%`, `%LocalAppData%`,
`%UserProfile%`, `%APPDATA%`, …) are expanded in any value.

### Logging

Diagnostic logging is **off by default**. Turn it on by setting `LogPath`
under `[Logging]` to a writable file. Missing parent directories are
created on first write.

```ini
[Logging]
LogPath=%LocalAppData%\TCOfficeView\host.log
```

Leave the value empty (or comment it out) to disable logging.

### Fallback Information Panel

Under `[FallbackUI]`:

- `FontFamily` — exact font name. Leave empty (the default) to auto-pick
  the first installed font from:
  Aptos Mono (Microsoft 365 / Office 2024) →
  Consolas (Windows Vista+) →
  Cascadia Mono (newer Windows) →
  Lucida Console (Windows 2000+) →
  Courier New (guaranteed final fallback).
- `FontSize` — font size in points; default `12`, clamped to `6..72`.
  The font is scaled to the actual monitor DPI when the panel is created.

```ini
[FallbackUI]
FontFamily=Cascadia Code
FontSize=13
```

## When MS Office Is Not Installed

If the operating system has no Preview Handler for the file's extension — or a
registered handler fails to load — the plugin renders a read-only text panel in
place of the preview. It states clearly that no preview handler was found and
lists what can still be read from the file system:

- file name and full path
- extension
- size (human-readable and exact byte count)
- created, modified and last-accessed timestamps

This way the Lister never shows an empty pane, and a user who installs
the plugin without Office still sees which file the Lister is on and
when it was last touched, along with instructions on how to enable real
previews.

## Troubleshooting

**Plugin does not activate.** Open **Configuration → Options → Plugins →
Lister plugins → Configure** and verify TCOfficeView is loaded and the
extension appears in its *Detect string*.

**Plugin shows the fallback panel instead of a real preview.** No Preview
Handler is registered for this file extension. The most common cause is
that MS Office is not installed; for `.msg` / `.vsdx` the relevant
component may simply not register a handler. Confirm by inspecting the
registry:

```
HKCR\.docx\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}
```

If that key is missing, install (or repair) Office.

**Plugin activates but the pane stays completely blank.** The handler
was created but rendered nothing. Enable diagnostic logging via the INI
to see where the load sequence stopped, and check Event Viewer for
crashes of `TCOfficeViewHost.exe`.

**TC crashes.** Should not happen thanks to process isolation. If it
does, check Event Viewer for crashes of `TCOfficeViewHost.exe`. The most
common cause is a 64/32-bit mismatch — for example a 64-bit Office where
only the 32-bit Preview Handler is registered. Try the other bitness of
the plugin.

---

# For Contributors

**The rest of this document is for developers who want to build, modify or extend the plugin.**

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

The script drives CMake for both bitnesses (x86 and x64) and packages the
artifacts into `dist\TCOfficeView.v<version>.zip` — the same auto-install
bundle described in *Installation*. The version is read from
`pluginst.inf` as the single source of truth; bump it there before
tagging a release.

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
