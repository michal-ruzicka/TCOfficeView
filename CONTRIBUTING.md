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

The same `build.cmd` is what the GitHub Actions CI workflow runs on every
push and pull request. The workflow uploads the resulting ZIP as a
downloadable artifact but does **not** publish GitHub Releases — that step
is done locally after GPG-signing the artifact (see *Release process*
below).

All actions in the workflow are pinned to a full commit SHA (required by
the repository's *Require actions to be pinned to a full-length commit SHA*
policy). Dependabot is configured to open monthly PRs that bump those SHAs
when upstream actions release new versions — do not update the SHAs by
hand.

## Repo Layout

- `src\` — C++ sources, INI/INF/manifest, `CMakeLists.txt`
- `build.cmd` — top-level build driver (the only file outside `src\`
  that the build touches)
- `build\` — CMake out-of-source build trees (gitignored)
- `dist\` — released ZIP bundles (gitignored)
- Markdown files (README, CONTRIBUTING, CHANGELOG, LICENSE, CLAUDE)
  live at the repo root and are bundled into every release ZIP.

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

### Universal File-Type Detection

The plugin ships with `detect=EXT="*"` in `pluginst.inf` — Total
Commander offers TCOfficeView every file the user previews.  Before
spawning the host process the plugin DLL runs `HasPreviewHandlerForExt`
(a registry walk mirroring the host's `FindPreviewHandlerClsid` lookup
chain — direct shellex → default ProgID → `OpenWithProgids` →
`SystemFileAssociations\<ext>` → `SystemFileAssociations\<PerceivedType>`)
and returns `nullptr` from `ListLoadW` / `LISTPLUGIN_ERROR` from
`ListLoadNextW` when no handler is registered.  TC then routes the file
to the next configured Lister plugin or its built-in viewer.

This is why there is no per-format table in `TCOfficeView.cpp`: format
support is whatever the user's machine has registered.  When adding a
new test file type, no plugin code change is needed — install the
relevant application and the next preview picks it up automatically.

The detection function in the plugin DLL is intentionally a near-copy
of the host's `FindPreviewHandlerClsid` rather than shared via a
header.  The two callers want different return types (`bool` vs.
`HRESULT + CLSID`) and the registry chain has been stable enough that
duplication costs less than a build-system split.  Keep them in
lock-step when editing — a mismatch would either spawn the host for
files it can't preview, or hand control back to TC for files we could.

### Plugin ↔ Host Wire Protocol

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

### Known Limitations

- **First-preview latency** – roughly 200–800 ms (cold start of the
  Office COM server). Subsequent previews are 100–300 ms.
- **Memory** — each host process holds ~30–80 MB while a preview is
  active. Successive previews in the same Lister session reuse the same
  host; cross-session pooling is not implemented.
- **`ListSendCommand` is a stub.** Print / find / copy from TC's Lister
  menu are not yet wired through to the host. Most preview handlers
  expose their own context menus inside the preview area, so this is
  rarely noticed in practice.

## Diagnostics

User-facing troubleshooting lives in [README.md](README.md#troubleshooting);
the developer-level notes live here.

**Host process names.** The plugin DLL spawns one of two helper
executables depending on bitness:

- 64-bit plugin → `TCOfficeViewHost.exe`
- 32-bit plugin → `TCOfficeViewHost_x86.exe`

Both run as `asInvoker`, host the preview handler in a COM STA, and
talk to the plugin DLL over a named pipe (`\\.\pipe\TCOfficeView_<pid>_<tick>`).
Look for these process names in Task Manager, Process Explorer or
Event Viewer when investigating a hang or a crash.

**Diagnostic log.** Off by default; enable in `TCOfficeView.ini`:

```ini
[Logging]
LogPath=%LocalAppData%\TCOfficeView\host.log
```

The log captures every step of the host's load sequence — registry
lookup → `CoCreateInstance` → `IInitializeWith*::Initialize` →
`IPreviewHandler::SetWindow` → `DoPreview` — with the HRESULT of each.
The first non-success HRESULT pinpoints the stage where things broke.

**Verifying a handler is registered.** The host's lookup chain is:

1. `HKCR\<ext>\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}`
2. `HKCR\<default ProgID of <ext>>\shellex\{8895b1c6-...}`
3. `HKCR\<each ProgID under <ext>\OpenWithProgids>\shellex\{8895b1c6-...}`
4. `HKCR\SystemFileAssociations\<ext>\shellex\{8895b1c6-...}`
5. `HKCR\SystemFileAssociations\<PerceivedType>\shellex\{8895b1c6-...}`

The plugin DLL's `HasPreviewHandlerForExt` walks the same chain.  If
none of these keys exist for the extension in question, no handler is
registered and the plugin will (correctly) decline the file.

**Event Viewer.** Faults in the host show up under *Windows Logs →
Application* as `Application Error` events naming the host EXE (most
useful for tracking down a misbehaving third-party preview handler).

**Bitness mismatches.** Both bitnesses ship in the auto-install ZIP.
Some preview handlers register only one bitness — switch to the other
plugin bitness via TC's plugin configuration if a particular file type
crashes consistently.

## Signing Policy

All commits merged into `main` must be signed. Tags `v*` must also be
signed. The repository's branch and tag protection rules enforce this
server-side — unsigned commits and tags are rejected on push.

Two distinct signing mechanisms are in use:

| What | Method | Key type |
|------|--------|----------|
| Git commits and annotated tags | SSH key signing | Your SSH key |
| Release ZIP files | GPG detached signature | Separate GPG key |

### Setting up SSH commit signing

You can reuse the same SSH key you already use to authenticate to GitHub.

**1. Register the key as a signing key on GitHub.**

Go to **Settings → SSH and GPG keys → New signing key** and paste your
public key. This is a separate entry from the authentication key, but
both entries can use the same key material.

**2. Configure Git for this repository.**

```
git config gpg.format ssh
git config user.signingkey "~/.ssh/id_ed25519.pub"
git config commit.gpgsign true
```

Replace `id_ed25519.pub` with your actual public key filename if
different. Omitting `--global` scopes these settings to this repository
only.

**3. Verify.**

```
git commit --allow-empty -m "test signing"
git log -1 --show-signature
```

Git should report `Good "git" signature` with your key fingerprint.

The same `gpg.format = ssh` setting is picked up by `git tag -s`, so
the release tagging step in *Release Process* below also uses your SSH
key automatically — no separate GPG key is needed for tags.

### GPG signing of release ZIPs

The distributable ZIP is signed with a GPG key (not the SSH key) to
produce a detached `.asc` signature that end users can verify without
having to trust GitHub's infrastructure. The signing key fingerprint is
`489C 5EC8 0FD6 2BE8 9E59  B4F7 19C1 3E8C E0F5 DB61` (available on
<https://keys.openpgp.org/>). The exact signing steps are in the
*Release Process* section below.

## Release Process

Releases are built and signed locally; no private key ever leaves the
developer's machine.

1. On a feature branch, bump `version=` in `src\pluginst.inf` and add a
   `## [X.Y.Z] – YYYY-MM-DD` entry to `CHANGELOG.md`.
2. Open a PR, get it reviewed, and merge into `main`.
3. On `main`, tag and push:
   ```
   git fetch && git checkout main && git pull
   git tag -s vX.Y.Z -m "Release X.Y.Z"
   git push origin vX.Y.Z
   ```
4. Run `build.cmd` locally to produce `dist\TCOfficeView.vX.Y.Z.zip`.
5. GPG-sign the ZIP:
   ```
   gpg --detach-sign --armor dist\TCOfficeView.vX.Y.Z.zip
   ```
   This creates `dist\TCOfficeView.vX.Y.Z.zip.asc`.
6. On the GitHub repository page, go to **Releases → Draft a new release**,
   select the `vX.Y.Z` tag, paste the CHANGELOG entry as the description,
   and attach both files (`.zip` and `.zip.asc`).

The CI workflow also produces the ZIP as a downloadable Actions artifact,
but that copy is unsigned and is intended for testing PRs only.

## License

This project is licensed under the Apache License 2.0 — see
[LICENSE.md](LICENSE.md) for the full text. Source files carry SPDX
identifiers (`SPDX-License-Identifier: Apache-2.0`). Release notes are in
[CHANGELOG.md](CHANGELOG.md).
