# CLAUDE.md

A pure-Perl salvager that extracts readable text from corrupt Microsoft and
OpenOffice files (`.doc`, `.xls`, `.ppt`, `.docx`, `.xlsx`, `.pptx`) when
the applications themselves refuse to open them. CLI with an optional Tk
GUI. Cross-platform — runs wherever Perl 5.10+ runs.

## Repo map

- `coffice2txt.pl` — main CLI / Tk GUI entry point.
- `coffice2txtworking.pl` — working copy / scratch script; treat as
  reference, not the shipped entry point.
- `xlscatPDP`, `ss2tk` — companion CLIs.
- `ParseExcel.pm`, `ReadPDP.pm`, `TempFolderCreator.pm`,
  `Utility2007PDP.pm`, `ZipMod2.pm` — top-level Perl modules used by the
  scripts above.
- `Spreadsheet-XLSX-TempFolderCreator/`,
  `Spreadsheet-XLSX-Utility2007PDP/`, `Spreadsheet-XLSX-ZipMod2/` —
  packaged versions of the same modules for CPAN-style installation.
- `cpanfile` — Perl dependency list.
- `test.pl` — test harness.
- `web/` — landing-page assets for the GitHub Pages site.
- `scripts/` — release packaging helpers.
- `VERSION`, `CHANGELOG.md` — version and change history.
- `.github/workflows/` — `build.yml` (CI), `pages.yml` (deploy to Pages
  on push to `main`), `release.yml` (build platform tarballs on `v*` tag).

## Branch policy

Work on the assigned feature branch:

1. Commit and push the feature branch.
2. **Open a PR from the feature branch to `main`** using the GitHub MCP
   tools (`mcp__github__create_pull_request`). Do not merge directly —
   the maintainer reviews and merges.
3. CI runs on the PR; Pages and Release pipelines fire only after merge
   to `main`.

## Releasing

- Push a `v*` tag to `main` to produce
  `coffice2txt-{windows,macos,linux,chromeos}-vX.Y.Z.{zip,tar.gz}`.
- Each archive ships with the platform's run script (`run.bat`,
  `run.command`, `run.sh`) and assumes the user has a system Perl + Tk.

## Verifying changes

- `perl test.pl` is the primary test harness.
- For format-specific changes, run the CLI against a known-corrupt fixture
  of the relevant format and confirm extracted text matches the prior
  baseline.
- If you change a top-level module (`ParseExcel.pm`, etc.) **also update
  the corresponding `Spreadsheet-XLSX-*/` package** — those are
  hand-mirrored and drift silently otherwise.

## Gotchas

- Tk is optional. Guard `use Tk;` so the CLI keeps working in headless /
  no-Tk environments (Linux containers, CI).
- The XLSX path uses bundled ZIP modules — do not `use Archive::Zip` or
  similar, the whole point is to handle archives those modules reject.
- Strict mode and `use warnings` are inconsistent across files (legacy
  code). Don't bulk-add them; fix per file when you have a test.
