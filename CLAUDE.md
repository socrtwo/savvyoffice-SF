# CLAUDE.md

Savvy Repair for Microsoft Office — recovers corrupt `.docx`, `.xlsx`, and
`.pptx` files using four recovery methods, entirely on-device. Two
implementations: a **cross-platform PWA under `web/`** (canonical, ships
to every release) and a **legacy VB.NET WinForms app** under `Savvy
Repair for Microsoft Office/` (Windows-only). The web build is canonical;
desktop launchers under `launchers/` just start a local static server and
open the browser.

## Repo map

- `web/` — the PWA. Implements all four recovery methods. **Edits go
  here** unless the task is specifically about the legacy VB.NET app.
- `Savvy Repair for Microsoft Office/`,
  `Savvy Repair for Microsoft Office.sln` — legacy VB.NET WinForms app.
  Source of the original four recovery methods, now reimplemented in
  `web/`.
- `launchers/` — tiny per-platform launcher binaries / scripts that start
  a local static server and open the browser. Used by desktop bundles.
- `releases/` — pre-packaged release archives committed to the repo.
- `.github/workflows/` — `build.yml` (CI), `pages.yml` (deploy `web/` to
  Pages on push to `main`), `release.yml` (build per-platform bundles on
  `v*` tag).

The four recovery methods (in `web/`):
1. **Zip structure repair** — `PK\x03\x04` scan + Immortal Inflater (the
   fault-tolerant DEFLATE from `socrtwo/Universal-File-Repair-Tool`).
2. **Strict XML validation** — truncate at first parse error, re-close
   open tags.
3. **Lax XML validation** — drop most-broken elements iteratively.
4. **Plain-text salvage** — pull `<w:t>` / `<a:t>` / `<t>` runs into a
   `.txt`.

## Branch policy

Work on the assigned feature branch:

1. Commit and push the feature branch.
2. **Open a PR from the feature branch to `main`** using the GitHub MCP
   tools (`mcp__github__create_pull_request`). Do not merge directly —
   the maintainer reviews and merges.
3. CI runs on the PR; Pages and Release pipelines fire from `main` only.

## Releasing

- Push a `v*` tag to `main` to produce per-platform bundles. Desktop
  bundles ship `web/` + the platform launcher from `launchers/`; mobile
  and ChromeOS install the PWA from the live page.

## Verifying changes

- PWA: serve `web/` locally and exercise all four methods against
  corrupt-`.docx` / `.xlsx` / `.pptx` fixtures.
- For Immortal Inflater changes, also confirm that truncated XML subfiles
  still yield partial bytes instead of crashing the unzipper — that's
  the whole reason this decoder exists.
- VB.NET app: open the `.sln` in Visual Studio. CI on `build.yml`
  validates this build.

## Gotchas

- The four methods are listed in increasing aggressiveness. Don't reorder
  them in the UI — users expect "method 1 first, last resort is
  method 4."
- The Immortal Inflater is a **never-throws** DEFLATE decoder. If you
  modify it, raising an exception on bad data is a bug, not a fix —
  return whatever bytes were successfully decoded.
- Lax XML validation can produce semantically-broken output. That's
  acceptable — the goal is recovery, not validity. Don't add a
  "validate again after lax repair" step.
- Desktop launchers are deliberately minimal — they exist only to open
  the browser to a local server. Don't grow them.
