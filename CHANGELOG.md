# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.1.2] - 2026-06-08

### Added

- `--sheet:length`, `--sheet:list`, and `--sheet:info` qualifier forms on
  `-s/--sheet` — inspect a workbook's sheet metadata without performing an
  extraction.
  - `--sheet:length` — prints the number of sheets as a bare number.
  - `--sheet:list` — prints one sheet name per line (markdown bullet list when
    routed to `md` or `pdf`).
  - `--sheet:info` — prints the count followed by a bullet list of names.
- `SheetQuery` type (`'length' | 'list' | 'info'`) added to the public type
  surface.
- `dispatchText` output function for routing pre-rendered plain text to all
  configured targets, used by sheet queries.

## [2.1.1] - 2026-06-05

### Fixed

- Test runner internal corrections and README clarifications.

## [2.1.0] - 2026-06-05

### Added

- `-x, --xml` option — maps a sheet to XML records; emits the result to the
  terminal or a file. An `--xml:root,row` qualifier overrides the default root
  and row element tag names.
- `--action:xml` output target for XML-formatted output.
- `--range sheet:"Name"` form — extracts the used range of a named sheet by
  quoting the sheet name after the colon.
- XML map files (`tests/maps/`) and additional fixture workbooks for XML tests.
- Expanded built-in `--test` runner: global, unit, and custom test modes with
  filter and replay support; substantially more coverage of output routing and
  format scenarios.

## [2.0.0] - 2026-05-31

### Added

- `--action:table` — renders output as an aligned, padded terminal table.
- `--action:csv` and `--action:markdown` output format modifiers.
- Extension-inferred targets: `--file report.pdf` routes to `pdf`; `--file
  report.md` routes to `md`. Any other extension stays a plain text file.
- `--assume-merge` switch — collapses repeated values from merged cells in
  `pdf` and `md` output (numeric duplicates are preserved).
- `-o, --orientation` option for PDF page layout (`landscape` / `portrait`).
- `-v, --version` flag — prints the installed version.
- Multi-workbook support — chain multiple `-b/--book` arguments, each with its
  own sheet selection and output routing, in a single command.
- `--action:var` named-capture form: `var=_name` and `file=out.txt` as
  positional key-value pairs within `--action`.
- Public API (`dist/index.js` + `dist/index.d.ts`) mirroring the CLI engine,
  so scripts and packages can import the extraction engine directly.

### Changed

- `pdf` and `md` output now uses a two-pass aligned table renderer: column
  widths are measured across all rows first, then every cell is padded to
  match. Empty columns are dropped; a blank data row stacks separate tables.
- Tokenizer extended to handle repeating `-b` arguments, output format
  modifiers on `--action`, and multi-target colon-joined action syntax.

## [1.0.0] - 2026-05-30

### Added

- Initial release.
- CLI: cell (`-c/--cell`), range (`-r/--range`), and title (`-t/--title`)
  extraction from `.xlsx` workbooks.
- A1-notation normalization: references are case-insensitive and ranges are
  auto-ordered (`C13:A2` equals `A2:C13`).
- Output routing to `stdout`, file, PDF, markdown, and shell variables
  (`--action:var`).
- Sheet selection with `-s/--sheet` — sticky; must immediately follow its
  workbook argument.
- `ee-capture` shell helpers for `bash`, `cmd`, and PowerShell to capture
  `--action:var` exports into shell variables.
- Built-in `--test` runner.
- `config.json` for runtime configuration (range-column separator, output
  defaults).

[unreleased]: https://github.com/isocialPractice/extract-excel/compare/v2.1.1...HEAD
[2.1.1]: https://github.com/isocialPractice/extract-excel/compare/v2.1.0...v2.1.1
[2.1.0]: https://github.com/isocialPractice/extract-excel/compare/v2.0.0...v2.1.0
[2.0.0]: https://github.com/isocialPractice/extract-excel/compare/v1.0.0...v2.0.0
[1.0.0]: https://github.com/isocialPractice/extract-excel/releases/tag/v1.0.0
