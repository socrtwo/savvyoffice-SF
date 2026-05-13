<!--MODERNIZED:v2-->
# Savvy Repair for Microsoft Office

[![Live page](https://img.shields.io/badge/live-app-ff2e93?style=for-the-badge)](https://socrtwo.github.io/savvyoffice-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/savvyoffice-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/savvyoffice-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/savvyoffice-SF?style=for-the-badge&color=22d3ee)](https://github.com/socrtwo/savvyoffice-SF/blob/main/LICENSE)
[![Last commit](https://img.shields.io/github/last-commit/socrtwo/savvyoffice-SF?style=for-the-badge&color=34d399)](https://github.com/socrtwo/savvyoffice-SF/commits)

Recover corrupt **`.docx`**, **`.xlsx`**, and **`.pptx`** files using four
recovery methods — entirely on your device. No upload. No server.

🌐 **Try it now:** <https://socrtwo.github.io/savvyoffice-SF/>
📦 **Downloads:** [Releases](https://github.com/socrtwo/savvyoffice-SF/releases)

---

## What it does

The four recovery methods from the original VB.NET app, re-implemented in
the browser:

1. **Zip structure repair** — scans the file byte-by-byte for `PK\x03\x04`
   local-file-header signatures and re-packs every recoverable entry.
   Uses a fault-tolerant DEFLATE decoder (the *Immortal Inflater* from
   [socrtwo/Universal-File-Repair-Tool](https://github.com/socrtwo/Universal-File-Repair-Tool))
   so that **truncated XML subfiles still yield partial bytes** instead
   of crashing the unzipper.
2. **Strict XML validation** — for every `.xml` / `.rels` part, truncate
   at the first parse error and re-close any still-open tags.
3. **Lax XML validation** — drop the most-broken element repeatedly
   until the part parses; recovers more data at the cost of some
   formatting.
4. **Plain-text salvage** — pull `<w:t>` / `<a:t>` / `<t>` runs into a
   `.txt`, the last-resort recovery when even the XML can't be fixed.

## Platforms

The web build is the canonical implementation. Native desktop installers
(built with Electron) ship the same web app in a real desktop window —
no browser or Python required. Mobile / ChromeOS users can install it as a PWA.

| Platform   | Native installer (recommended)                                                                                         | Portable bundle (no install)                                                           |
| ---------- | ---------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| 🌐 Web      | Open <https://socrtwo.github.io/savvyoffice-SF/> — no install needed.                                                  | —                                                                                      |
| 🪟 Windows  | `SavvyRepair-*-win-setup.exe` — NSIS installer (Next → Next → Finish)                                                 | `SavvyRepair-windows-*.zip` — unzip, run **Run Savvy Repair.cmd** (needs Python)       |
| 🍎 macOS    | `SavvyRepair-*-mac.dmg` — drag to Applications; right-click → Open on first launch                                    | `SavvyRepair-macos-*.zip` — unzip, right-click **Run Savvy Repair.command** → Open    |
| 🐧 Linux    | `SavvyRepair-*-linux.AppImage` — `chmod +x` then run; or `*_amd64.deb` for apt                                       | `SavvyRepair-linux-*.tar.gz` — extract, run `./run-savvy-repair.sh`                   |
| 💻 ChromeOS | Open the web app, then **menu → Install app** (or the install icon in the address bar).                                | —                                                                                      |
| 🤖 Android  | Open the web app in Chrome, then **menu → Add to Home Screen / Install app**.                                          | —                                                                                      |
| 📱 iOS      | Open the web app in Safari, tap **Share → Add to Home Screen**.                                                        | —                                                                                      |

Once installed as a PWA, the app works fully offline.

## Privacy

Everything runs locally in your browser. Your document is never uploaded
anywhere — there is no server-side component. You can verify this by
disconnecting from the network after the page loads.

## How it works under the hood

* **`web/index.html`** — UI shell, PWA manifest, drag-and-drop picker.
* **`web/immortal-inflate.js`** — fault-tolerant DEFLATE decoder and a
  PK-signature scanner that survives truncated streams. Lets us recover
  entries from archives that JSZip alone refuses to open.
* **`web/app.js`** — orchestrates the four repair methods, packages
  results back into clean archives via JSZip.
* **`web/sw.js`** — service worker for offline use.

## Source heritage

This started life on SourceForge in 2014 as a VB.NET WinForms app that
shelled out to `7z.exe`, `xmllint.exe`, `xmlval.exe`, `trunc.exe`, and
`doctotext.exe`. That code is still in this repo under
[`Savvy Repair for Microsoft Office/`](./Savvy%20Repair%20for%20Microsoft%20Office/)
for historical reference; the modern, cross-platform release is the web
build in [`web/`](./web/). Both are MIT licensed.

The legacy SourceForge project page is at
<https://sourceforge.net/projects/savvyoffice/>.

## Building from source

### Web / PWA (all platforms)

No build step — `web/` is plain HTML/JS/CSS. To run locally:

```bash
cd web && python3 -m http.server 8765
# open http://localhost:8765/
```

### Electron desktop app (Windows / macOS / Linux)

Requires Node.js 20+.

```bash
npm install               # install Electron + electron-builder
npm start                 # launch in dev mode

npm run dist:win          # build Windows NSIS .exe + portable .exe → dist-electron/
npm run dist:mac          # build macOS .dmg (x64 + arm64)          → dist-electron/
npm run dist:linux        # build Linux .AppImage + .deb             → dist-electron/
```

CI builds all three in parallel via `.github/workflows/release.yml` when a
`v*` tag is pushed.

### Original VB.NET app (Windows only)

Open `Savvy Repair for Microsoft Office.sln` in Visual Studio 2017+ and
build. The included `.github/workflows/build.yml` builds it on
`windows-latest` with MSBuild.

## Contributing

Issues and pull requests welcome at
<https://github.com/socrtwo/savvyoffice-SF/issues>.

## License

MIT — see [LICENSE](LICENSE).

The fault-tolerant inflater is adapted from
[socrtwo/Universal-File-Repair-Tool](https://github.com/socrtwo/Universal-File-Repair-Tool)
(also MIT).
