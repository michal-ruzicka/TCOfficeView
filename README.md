# TC Office Lister Plugin

A Total Commander Lister plugin for previewing MS Office documents
(DOCX, XLSX, PPTX, …) through the native **Windows Preview Handlers**
registered by Office itself.

## Architecture

```
┌──────────────────────────────────────────────────────┐
│  Total Commander (F3 / Quick View)                   │
│                                                      │
│  ┌────────────────────────────────────────────┐      │
│  │  tcoffice.wlx(64)   ← TC loads this DLL    │      │
│  │                                            │      │
│  │  • ListLoad()    → creates child HWND      │      │
│  │  • LaunchHost()  → spawns host process     │      │
│  │  • named-pipe transport (LOAD/RESIZE/CLOSE)│      │
│  └────────────────────────────────────────────┘      │
└──────────────────────────────────────────────────────┘
                  │  named pipe (UTF-16, message mode)
                  ▼
┌──────────────────────────────────────────────────────┐
│  officehost.exe  (isolated STA process)              │
│                                                      │
│  • CoCreateInstance(Word/Excel/PowerPoint           │
│                     Preview Handler CLSID)           │
│  • IInitializeWithFile/Stream::Initialize(path)      │
│  • IPreviewHandler::SetWindow(hwnd) + DoPreview()    │
│                                                      │
│  → Office renders directly into TC's HWND            │
└──────────────────────────────────────────────────────┘
```

Both components are pure C++ / Win32 / COM. No .NET, no managed runtime,
no extra dependencies beyond the Visual C++ static runtime that is linked
into the binaries themselves.

## Why an isolated host process?

1. **Stability.** Office preview handlers occasionally crash. A crash in the
   host does not bring down Total Commander. (Windows Explorer follows the
   same pattern with its own `prevhost.exe`.)
2. **STA threading.** COM preview handlers require a single-threaded
   apartment. Guaranteeing STA inside a DLL loaded by an arbitrary host
   process (TC) is fragile; an own EXE makes the apartment explicit.
3. **Bitness flexibility.** Office preview handlers are registered per
   bitness. 64-bit TC paired with 32-bit Office (or vice versa) is handled
   by COM's cross-process activation: the host calls `CoCreateInstance`
   with `CLSCTX_LOCAL_SERVER`, and Windows brings up the handler in the
   matching bitness.

## Build

### Prerequisites

- Visual Studio 2022 Build Tools (workload: *Desktop development with C++*)
- CMake 3.20+

That is the entire list — no .NET SDK is required.

### Build steps

```cmd
build.cmd
```

Artifacts are written to `dist\`:

- `tcoffice.wlx`        — plugin DLL for 32-bit Total Commander
- `tcoffice.wlx64`      — plugin DLL for 64-bit Total Commander
- `officehost.exe`      — host process used by the 64-bit plugin
- `officehost_x86.exe`  — host process used by the 32-bit plugin

## Installation

1. Copy the contents of `dist\` into a persistent folder, e.g. `C:\Tools\TCOffice\`.
2. In Total Commander: **Configuration → Options → Plugins → Lister plugins → Configure**.
3. Click **Add** and pick `tcoffice.wlx` (32-bit TC) or `tcoffice.wlx64` (64-bit TC).
4. TC registers the plugin for the extensions declared in `DetectString`.

Test it with **F3** on any `.docx` / `.xlsx` / `.pptx` file.

## Supported formats

Depends on which Preview Handlers are installed on the system. With Office
installed you typically get:

| Application | Extensions |
|-------------|------------|
| Word        | DOC, DOCX, DOCM, RTF |
| Excel       | XLS, XLSX, XLSM, XLSB |
| PowerPoint  | PPT, PPTX, PPTM |
| Visio       | VSD, VSDX |
| Outlook     | MSG |

The plugin advertises support for the union of these extensions; if a
specific handler is not installed, the host falls back gracefully and TC
will use its built-in viewer instead.

## Plugin ↔ host wire protocol

Text-based, UTF-16 LE, lines terminated by `\n`, transported over a
message-mode named pipe.

**Plugin → host:**

```
LOAD <absolute-path>
RESIZE <width> <height>
CLOSE
```

**Host → plugin:**

```
OK
ERR <message>
```

`RESIZE` is fire-and-forget — the host does not acknowledge it, because the
plugin coalesces `WM_SIZE` events through a 50 ms timer and a steady stream
of resize messages would otherwise saturate the pipe during a drag.

## Known limitations

- **First-preview latency.** The first preview after TC starts takes
  roughly 200–800 ms (cold start of the Office COM server). Subsequent
  previews are 100–300 ms. Switching to a native host removed the ~150–
  300 ms .NET startup cost the earlier prototype had.
- **Memory.** Each host process holds ~30–80 MB while a preview is active.
  Successive previews inside the same Lister session reuse the host
  (`ListLoadNextW` sends a new `LOAD` to the running host); cross-session
  pooling is not implemented yet.
- **`ListSendCommand`** is currently a stub. Print / find / copy from TC's
  Lister menu are not yet wired through to the host. Most preview handlers
  expose their own context menus inside the preview area, so this is rarely
  noticed in practice.

## Troubleshooting

**Plugin does not activate.** Verify in the TC Lister Plugins dialog that
the plugin is loaded and that the extension appears in the *Detect string*.

**Plugin activates but the window stays empty.** A Preview Handler is most
likely not registered for the given type. Confirm in the registry:

```
HKCR\.docx\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}
```

**TC crashes.** Should not happen thanks to process isolation. If it does,
check Event Viewer for crashes of `officehost.exe`. The most common cause
is a 64/32-bit mismatch — for example a 64-bit Office where only the
32-bit Preview Handler is registered. Try the other bitness of the plugin.
