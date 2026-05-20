# TCOfficeView: a Total Commander Lister Plugin that Previews Microsoft Office Documents

[![Build](https://github.com/michal-ruzicka/TCOfficeView/actions/workflows/build.yml/badge.svg)](https://github.com/michal-ruzicka/TCOfficeView/actions/workflows/build.yml)

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

**Project page:** <https://github.com/michal-ruzicka/TCOfficeView> — source
code, releases and issue tracker.

**Binary releases:** <https://github.com/michal-ruzicka/TCOfficeView/releases>

---

## Installation

The recommended way is the **Total Commander auto-installer:**

1. **Download** `TCOfficeView.vX.Y.Z.zip` from the [**latest release**](https://github.com/michal-ruzicka/TCOfficeView/releases).
2. In Total Commander, **navigate to the ZIP** file and press **Enter**.
3. TC detects the bundled `pluginst.inf` and asks whether to install the
   plugin — **confirm**.
4. Press **F3** or **Ctrl+Q** on a `.docx` / `.xlsx` / `.pptx` file to verify.

If you prefer to install manually, unzip the archive into a persistent
folder (for example `C:\Tools\TCOfficeView\`) and add `TCOfficeView.wlx`
(32-bit TC) or `TCOfficeView.wlx64` (64-bit TC) under **Configuration →
Options → Plugins → Lister plugins → Configure → Add**.

### Verifying Releases

Each release ZIP is accompanied by a detached GPG signature file
(`TCOfficeView.vX.Y.Z.zip.asc`). Before installing, verify that the
archive has not been tampered with:

```
gpg --keyserver keys.openpgp.org --recv-keys 489C5EC80FD62BE89E59B4F719C13E8CE0F5DB61
gpg --verify TCOfficeView.vX.Y.Z.zip.asc TCOfficeView.vX.Y.Z.zip
```

GPG should report `Good signature from "Michal Růžička <ruzicka.mich@gmail.com>"`.
The full fingerprint of the signing key is `489C 5EC8 0FD6 2BE8 9E59  B4F7 19C1 3E8C E0F5 DB61`.

## Usage

There is no UI to configure at runtime — just press **F3** (or use Quick View
[**Ctrl+Q**]) on any supported file. The preview tracks the Lister pane size and
selecting a different file in the panel loads it into the same Lister session.

### Supported Formats

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

> **Quick previews use a simplified rendering pipeline.** Office's
> built-in preview components — used by Windows Explorer's preview
> pane and by this plugin in the default mode — render through a
> reduced pipeline optimised for speed and stability inside a host
> window, not the full editing UI. For Word this means a flowing Web
> Layout without page breaks, headers or footers; Excel shows a
> simplified grid; PowerPoint shows static slides without transitions.
> An opt-in **Full mode** that launches the real Word, Excel or
> PowerPoint application and embeds its window into the Lister pane
> is available; see [Application Render Mode](#application-render-mode)
> below.

> **MSG and VSDX caveat.** Recent Office and Outlook installs sometimes
> ship without preview support for `.msg` or `.vsdx` — the New Outlook
> in particular has dropped the classic MSG preview. In those cases
> Windows Explorer's own preview pane is empty too, and the plugin
> falls back to the information panel.

### Configuration

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

#### Logging

Diagnostic logging is **off by default**. Turn it on by setting `LogPath`
under `[Logging]` to a writable file. Missing parent directories are
created on first write.

```ini
[Logging]
LogPath=%LocalAppData%\TCOfficeView\host.log
```

Leave the value empty (or comment it out) to disable logging.

#### Fallback Information Panel

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

#### Application Render Mode

By default all previews use the **quick mode** — the built-in Office
preview component that the plugin hosts inside the Lister pane. This
is fast (~200–800 ms to first display), memory-light (~30–80 MB) and
stable, but the rendering pipeline is simplified (Word in Web Layout
without page breaks, Excel in a simplified grid, PowerPoint slides
without transitions).

Setting an application to `full` switches to a slower but visually
faithful mode: the real Word, Excel or PowerPoint application is
launched in the background, opens the file read-only, and its main
window is embedded into the Lister pane. All three applications can
be configured independently.

```ini
[Mode]
Word=full
Excel=full
PowerPoint=full
```

A small **mode-switch button** in the top-right of every Word /
Excel / PowerPoint preview lets you flip the current preview between
quick and full on demand — useful when you want a quick glance at one
file in full mode without changing your default, or vice versa. The
switch only affects the currently displayed file; selecting another
file (or re-opening the same one) goes back to your INI default. The
button is hidden for file types that have no full-mode equivalent
(`.msg`, `.vsdx`).

What full mode does for each application:

- **Word** — shows the document in Print Layout (page boundaries,
  headers, footers, page numbers) with the page scaled to the Lister
  pane width. The zoom re-fits automatically when you resize the
  pane. Rulers are hidden and the preview is truly read-only (typing
  in the document is blocked).
- **Excel** — opens the workbook read-only with the zoom set to
  100%. Excel's initial layout fills the Lister pane on load.
  **Known limit:** Excel does not relayout when the Lister pane is
  later resized — the content stays anchored to its initial area.
  Close the Lister (Esc) and reopen it with F3/Ctrl+Q to get a
  fresh layout at the new pane size.
- **PowerPoint** — opens the presentation read-only with the slide
  scaled to fit the Lister pane. The zoom re-fits automatically when
  you resize the pane. PowerPoint's main window appears on screen for
  a fraction of a second before being embedded into the Lister (a
  brief visible flash); Word and Excel embed silently.

Full mode tradeoffs to be aware of:

- **Cold start ~2–4 s** the first time an Office document of a given
  application is opened. Subsequent documents of the *same*
  application within the *same* Lister window load faster (~0.5–1 s)
  because the running Office instance is reused. Switching between
  file types (`.docx` → `.xlsx`) quits the previous application and
  spins up the next one, so that switch pays the cold-start cost
  again.
- **Only one application embedded at a time.** The Lister pane is a
  single embed point — you cannot have a Word document and an Excel
  workbook visible in the same Lister.
- **Memory ~100–300 MB** per running Office instance.
- **Requires a full Microsoft Office installation** of the relevant
  application — not Office Viewer, not LibreOffice.
- **Falls back to quick mode** automatically on any failure (Office
  missing, document password-protected, …), so you always get
  *some* preview.
- **No global Office settings are changed.** Full mode only touches
  settings scoped to the open preview window — your standalone Word,
  Excel and PowerPoint will start up exactly as you left them. The
  ribbon, for example, stays visible in the preview because hiding
  it would hijack your global Office settings.

### When MS Office Is Not Installed

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

### Troubleshooting

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

## Contributing

Build instructions, repo layout and architecture notes for developers live
in [CONTRIBUTING.md](CONTRIBUTING.md).

## License

This project is licensed under the Apache License 2.0 — see
[LICENSE.md](LICENSE.md) for the full text. Source files carry SPDX
identifiers (`SPDX-License-Identifier: Apache-2.0`). Release notes are in
[CHANGELOG.md](CHANGELOG.md).
