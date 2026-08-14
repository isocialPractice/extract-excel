# `extract-excel` — TODO

Tracking for the `extract-excel` CLI + API. Checked items are implemented and
covered by `extract-excel --test`.

## Done

- [x] Bespoke CLI tokenizer (default book, repeating `-b`, sticky `-s/--sheet`).
- [x] Cell / range / title extraction with A1 normalization.
- [x] Multi-workbook, multi-destination routing in one command.
- [x] Output targets: `stdout`, `file`, `pdf`, `md`, `var`.
- [x] Output formats: `text`, `table`, `csv`, `markdown` (+ extension inference).
- [x] `--range sheet` / `-r sheet:"Name"` used-range preset.
- [x] `pdf`/`md` two-pass aligned table (measure column widths, then pad). A
      blank row stacks a sheet into separate aligned tables; empty columns drop.
- [x] `--file out.pdf` / `--file out.md` imply the `pdf` / `md` target by
      extension (a non-`.md`/`.pdf` final extension stays a plain text file).
- [x] `--assume-merge` collapses repeated text from spanned merges in `pdf`/`md`
      output (numeric duplicates kept); sticky/positional per book.
- [x] `-o/--orientation` for PDF page layout (`landscape` / `portrait`).
- [x] `--action:var` capture helpers for bash / cmd / PowerShell.
- [x] Built-in `--test` runner (global / unit / custom, with filters + replay).
- [x] `-v/--version` flag.
- [x] Public API mirroring the CLI engine, shipped with `.d.ts` types.

## Next

### Output fidelity

- [ ] True colspan/rowspan in PDF (draw a real table grid instead of repeating
      the master value across spanned cells).
- [ ] Preserve Excel number/date formatting (currently numbers stringify and
      dates render ISO-8601).
- [ ] Configurable CSV dialect (delimiter, quoting, line endings).
- [ ] `--action:var` multi-line values (currently joined with `;` to stay
      assignable) — optional base64 / array modes.

### Extraction

- [ ] Whole-column / whole-row ranges (`A:A`, `2:2`).
- [ ] Multiple title columns extracted as one aligned block.
- [ ] Formula evaluation (currently reads the cached result value).
- [ ] `--title` matching options (exact / contains / regex).

### Workbook support

- [ ] `.xlsm` macro and `.xlsb` binary workbook reading.
- [ ] Password-protected workbooks.
- [ ] Streaming reader for very large workbooks (avoid loading the full grid).

### CLI / DX

- [ ] Per-option specific help (`--help range`, `--help action`).
- [ ] Shell completion scripts.
- [ ] Stricter validation errors with suggested corrections.

### Testing

- [ ] More fixture-based assertions against the bundled `.xlsx` files.
- [ ] Snapshot tests for `table` / `markdown` / `pdf` rendering.
- [ ] CI workflow (build + `--test`) under `.github/workflows`.
