# Changelog

All notable changes to this project will be documented in this file.

## [0.2.0] - 2026-08-30

### Security

- SheetJS dependency moved from npm `xlsx@0.18.5` to the official CDN tarball `xlsx@0.20.3` (<https://cdn.sheetjs.com>). npm's 0.18.5 carries CVE-2023-30533 (prototype pollution via crafted workbooks, fixed upstream in 0.19.3) and CVE-2024-22363 (ReDoS, fixed upstream in 0.20.2); fixed releases are only distributed through the official CDN. All Excel tool tests pass against 0.20.3.
- The five Office libraries (`docx`, `jszip`, `mammoth`, `pptxgenjs`, `xlsx`) are build-time-only and moved to `devDependencies`: they are inlined into `lib/index.js` by esbuild and never resolved at runtime, so installs of the published plugin fetch nothing beyond the plugin tarball (no cdn.sheetjs.com access required downstream). The host bundle shrinks from 4.0 MB to ~3.2 MB.

### Added

- `enablePptTools` config switch (default `true`) so this plugin can coexist with dedicated presentation plugins such as dsh-ppt, which register a colliding `ppt_create` that DSH rejects at startup. With `enablePptTools: false` only the five Word/Excel tools are registered. Declared through a schemastery `Config` validated by the Loader.

## [0.1.0] - 2026-08-15

### Added

- `word_create` / `word_read` for Word `.docx` documents.
- `excel_create` / `excel_read` / `excel_update` for Excel `.xlsx` workbooks.
- `ppt_create` / `ppt_read` for PowerPoint `.pptx` decks.
- PNG/JPG/GIF image embedding in `ppt_create` with explicit or automatic placement.
- Per-slide image count reporting in `ppt_read`.
- Workspace confinement (`session.header.cwd` + realpath check), atomic writes, overwrite protection, and size/cell limits.
- Unit/integration tests and GitHub Actions CI.
