# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.4] - 2026-05-09

### Added
- Multi-platform release bundles (Windows, macOS, Linux, ChromeOS, Android, iOS, Web).
- Platform-specific launcher scripts: `run.bat`, `run.command`, `run.sh`,
  `setup-termux.sh`, `setup-ish.sh`.
- GitHub Actions release workflow that builds and uploads all bundles on tag push.
- `cpanfile` declaring runtime dependencies for `cpanm --installdeps .`.
- `CHANGELOG.md` and `VERSION` files.
- Modernized README with per-platform installation and quick-start sections.

### Changed
- `.gitignore` updated for modern Perl/CPAN tooling and editor artifacts.
- Existing Pages workflow continues to publish the README-driven landing page.

### Notes
- This release expands platform coverage beyond v1.0.3 (which shipped a
  Windows-only installer). The Windows bundle here is a source-based zip that
  uses Strawberry Perl rather than a packaged `.exe`; users wanting the
  one-click installer can still download it from the v1.0.3 release page.

## [1.0.3] - 2026-04-06

### Added
- Windows installer (`corrupt_office_salvager_setup_1.0.3_without_adware.exe`)
  published as the first GitHub release after the SourceForge migration.
