# Changelog

All notable changes to this project are documented in this file.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-05-09

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
- This is the first tagged release after the SourceForge → GitHub migration.
  Code is unchanged from prior history; the release packages the existing
  scripts with platform-aware launchers and documentation.
