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
- [**CMake**](https://cmake.org/) 3.20 or newer. Use CMake bundled with Visual 
  Studio or install separately from cmake.org, via `winget install Kitware.CMake`, 
  or via `choco upgrade cmake`. The build uses the NMake Makefiles generator, 
  which has no cmake version requirement beyond the minimum above.

No .NET SDK, no extra runtime — the project is pure C++ / Win32 / COM,
linked against the static MSVC runtime.

## Repo Layout

| Path | Description |
|---|---|
| `.github/` | GitHub Actions CI workflow (`workflows/build.yml`) |
| `build\` | CMake out-of-source build trees (gitignored) |
| `dist\` | Released ZIP bundles (gitignored) |
| `src\` | C++ sources, INI/INF/manifest, `CMakeLists.txt` |
| `.gitattributes` | Line-ending normalization rules |
| `.gitignore` | Excludes `build\` and `dist\` from version control |
| `.vsconfig` | Visual Studio component selection; pins the MSVC toolset and Windows SDK versions for the CI workflow and `setup-build-environment-example.cmd` |
| `build.cmd` | Top-level build driver — the only file outside `src\` that the build touches |
| `CHANGELOG.md` | Release notes in Keep a Changelog format; bundled in every release ZIP |
| `CLAUDE.md` | Project notes for Claude Code; bundled in every release ZIP |
| `CONTRIBUTING.md` | Developer documentation (this file); bundled in every release ZIP |
| `LICENSE.md` | Apache License 2.0; bundled in every release ZIP |
| `README.md` | End-user documentation; bundled in every release ZIP |
| `setup-build-environment-example.cmd` | One-shot script: downloads sources from GitHub, installs Visual Studio Build Tools and CMake, and runs `build.cmd`; designed for VMs and Windows Sandbox |

## Build

From any command prompt where `vswhere.exe` can be found:

```cmd
build.cmd
```

The script locates Visual Studio Build Tools via `vswhere` and activates the
compiler environment itself — no Developer Command Prompt is needed. It drives
CMake (NMake Makefiles generator) for both bitnesses (x86 and x64) using the
sources under `src\` and packages the artifacts into
`dist\TCOfficeView.v<version>.zip` — the same auto-install bundle described
in *Installation*. The version is read from `src\pluginst.inf` as the single
source of truth; bump it there before tagging a release.

To set up a build environment from scratch on a fresh Windows installation
(VM, [Windows Sandbox](https://learn.microsoft.com/windows/security/application-security/application-isolation/windows-sandbox/),
or similar), `setup-build-environment-example.cmd` in the repo root downloads 
the sources, installs the required toolchain, and runs `build.cmd` 
automatically. Open it in a text editor before running — it is intentionally 
readable so you can see and trust every step.

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

### Reproducible Builds

Every build from the same source commit produces byte-identical binaries,
regardless of when or where the build runs. Users can therefore independently
verify that a published ZIP was built from the published source code without
having to trust the build infrastructure.

#### What makes the build deterministic

Four independent sources of non-determinism are eliminated:

| Source | Mechanism |
|---|---|
| PE `TimeDateStamp` field in `.dll` / `.exe` | Linker flag `/BREPRO` zeros it |
| Non-deterministic hash in MSVC `.obj` files | Compiler flags `/Brepro /experimental:deterministic` |
| Absolute source paths in debug info | Compiler flag `/pathmap:<src-dir>=./` strips the machine-local prefix |
| File modification times in ZIP entries | Normalized to `release-date=` read from `src\pluginst.inf` |

Three additional parameters are pinned so that "same source" also means "same
compiler output":

| Parameter | Pinned to | Files that must match |
|---|---|---|
| Exact `cl.exe` version | `19.51.36246` | `build.cmd` (`EXPECTED_CL`) |
| MSVC toolset (major.minor) | `14.51` | `build.cmd` (`CMAKE_T`), `.vsconfig`, CI workflow |
| Windows SDK | `10.0.26100.0` | `build.cmd` (`CMAKE_SDK`), `.vsconfig`, CI workflow |

`build.cmd` checks `cl.exe` against `EXPECTED_CL` and prints a warning if
they differ — the build continues regardless. The check exists so that when
a local and a CI build produce different hashes, the first thing to look at
in both logs is the `cl.exe` version line. Two builds from the same toolset
major.minor but different `cl.exe` patch versions will produce different
binaries; this happens when VS patches land at different times on local
machines and CI runner images. Update `EXPECTED_CL` (with a matching VS
update on both sides) when you want to re-establish bit-for-bit parity.

The toolset is selected by passing `-vcvars_ver=%CMAKE_T%` to `VsDevCmd.bat`
before invoking CMake. The SDK is enforced by CMake:

```
cmake ... -DCMAKE_SYSTEM_VERSION=10.0.26100.0
```

If the SDK is missing, CMake fails with a descriptive error rather than
silently falling back to a different version.

#### Verifying a build

The independent reference point is the SHA-256 of the ZIP itself, not files
inside it — `sha256sums.sha256` is bundled inside the same archive, so it is only
as trustworthy as the ZIP it came from.

To verify a release:

1. Run `build.cmd` locally on the release commit. The script prints the SHA-256
   of the produced ZIP at the end of its output.
2. Compare that hash against the SHA-256 printed in the *Build (x86 + x64)*
   step of the [GitHub Actions run](https://github.com/michal-ruzicka/TCOfficeView/actions)
   for the same commit.
3. Optionally, compare against the published release ZIP's SHA-256:
   ```
   certutil -hashfile dist\TCOfficeView.vX.Y.Z.zip SHA256
   ```

If all three match, the released ZIP is identical to what the source builds.

`sha256sums.sha256` inside the ZIP is useful as a convenience check after
extraction — to confirm that an individual file was not corrupted during
download or extraction — but it is not an independent trust anchor.

#### Upgrading the pinned toolchain

When you adopt a new MSVC toolset or Windows SDK version, update **all three
locations** listed below to keep them in sync. A mismatch causes either a
build failure (CMake cannot find the requested toolset) or silent divergence
between local and CI builds.

##### MSVC toolset

Identify the installed toolset version:

```
cl.exe /Bv
```

The path in the output contains the toolset directory, e.g.:
`...\VC\Tools\MSVC\`**`14.51`**`.36231\bin\...`

Use the first two numbers (`14.XX`) as the toolset token and the full
four-part version (e.g. `19.51.36244`) as the exact compiler version.
Update:

1. **`build.cmd`** — the three pinned variables at the top of the file:
   ```
   set EXPECTED_CL=19.51.XXXXX
   set CMAKE_T=14.XX
   set CMAKE_SDK=10.0.XXXXX.X
   ```
2. **`.vsconfig`** — the `VC.Tools` component entry (this also covers the
   CI workflow, which installs the toolchain from `.vsconfig`):
   ```json
   "Microsoft.VisualStudio.Component.VC.14.XX.x86.x64"
   ```

##### Windows SDK

The SDK version is a five-digit build number, visible in VS Installer under
the installed *Windows 11 SDK* component (e.g. `10.0.`**`26100`**`.0`).
Update:

1. **`build.cmd`** — both `cmake` configure lines:
   ```
   -DCMAKE_SYSTEM_VERSION=10.0.NNNNN.0
   ```
2. **`.vsconfig`** — the SDK component entry (this also covers the
   CI workflow, which installs the toolchain from `.vsconfig`):
   ```json
   "Microsoft.VisualStudio.Component.Windows11SDK.NNNNN"
   ```

##### Visual Studio major version

Each major Visual Studio release increments the product version number in the
installation path (`\17\` for VS 2022, `\18\` for VS 2026, …). The build
uses the NMake Makefiles generator driven by `VsDevCmd.bat`, so no CMake
generator flag needs updating. What does change:

1. **`.vsconfig`** — the toolset and SDK component IDs change with each major
   VS release; update those using the guidance above.
2. **`.github/workflows/build.yml`** — update the VS installer URL if the
   stable channel URL changes (currently `https://aka.ms/vs/18/stable/vs_buildtools.exe`).

Ensure CMake is new enough to support any `CMakeLists.txt` features in use
(`winget install Kitware.CMake` or `choco upgrade cmake`).

##### After any toolchain change

Delete the stale CMake build trees so CMake picks up the new selection on the
next run:

```
rmdir /S /Q build\x64 build\x86
build.cmd
```

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
