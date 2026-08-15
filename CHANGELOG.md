# Changelog

All notable changes to this project will be documented in this file.

## [0.1.0] - 2026-08-15

### Added

- `word_create` / `word_read` for Word `.docx` documents.
- `excel_create` / `excel_read` / `excel_update` for Excel `.xlsx` workbooks.
- `ppt_create` / `ppt_read` for PowerPoint `.pptx` decks.
- PNG/JPG/GIF image embedding in `ppt_create` with explicit or automatic placement.
- Per-slide image count reporting in `ppt_read`.
- Workspace confinement (`session.header.cwd` + realpath check), atomic writes, overwrite protection, and size/cell limits.
- Unit/integration tests and GitHub Actions CI.
