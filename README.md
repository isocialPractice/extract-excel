# extract-excel

Command-line tool **and** API to extract Microsoft Excel (`.xlsx`) and CSV data.
Pull individual cells, ranges, or whole columns by header title, then route the
output to the terminal, files, PDFs, or shell variables — across one or many
workbooks in a single command.

- **CLI**, designed to be installed globally via `npm link`.
- **API**, so scripts, frameworks, and other npm projects can reuse the engine.

---

## Installation

```bash
npm install
npm run build
npm link        # exposes the global `extract-excel` command
```

Or add it as a dependency and import the API:

```bash
npm install extract-excel
```

---

## CLI usage

```bash
extract-excel [file|app] [option] [task] [arguments...]
```

The first parameter is either an **app option** (`--help`, `--test`) or a
**workbook**. When it is not an app option it is treated as `-b/--book` by
default, so `extract-excel book.xlsx ...` just works.

### Options

| Option              | Type           | Description                                               |
| :------------------ | :------------- | :-------------------------------------------------------- |
| `-a, --action`      | output:multi   | Direct output (`stdout`, `file`, `pdf`, `md`, `var`).     |
| `-b, --book`        | file:single    | Workbook to extract data from.                            |
| `-c, --cell`        | extract:single | Extract a single cell, e.g. `-c A2`.                      |
| `-f, --file`        | output:single  | Write output to a file (`.pdf`/`.md` route by extension). |
| `-h, --help`        | app:doc        | Show global (`--help`) or specific help.                  |
| `-o, --orientation` | switch:string  | PDF page orientation (`landscape` \| `portrait`).         |
| `-r, --range`       | extract:multi  | Extract a range, e.g. `-r A2:F45`.                        |
| `-s, --sheet[:q]`   | select:single  | Select a sheet (must follow its workbook). Qualifier form: `:length` — sheet count; `:list` — sheet names; `:info` — count + names. |
| `-t, --title`       | extract:multi  | Extract a column by its header title.                     |
| `-v, --version`     | app:doc        | Print the installed version.                              |
| `-x, --xml`         | output:single  | Output the sheet's mapped table as XML.                   |
| `--assume-merge`    | switch:bool    | Collapse spanned merge cells in `pdf`/`md` output.        |
| `--test`            | app:test       | Run `global`, `unit`, or `custom` tests.                  |

### Rules worth knowing

- **`-s/--sheet` is sticky.** It is only honored when it *immediately* follows
  its workbook. `-b book.xlsx --cell A1 --sheet Data` silently ignores the
  sheet, because `--cell` came first.
- **Sheet queries skip extraction.** `--sheet:length`, `--sheet:list`, and
  `--sheet:info` are qualifier forms that inspect the workbook instead of
  selecting a sheet for extraction (see [Sheet queries](#sheet-queries) below).
- **Ambiguous sheets error out.** A workbook with more than one sheet and no
  `--sheet` raises a clear error listing the available sheets.
- **References are normalized.** `yc30`, `r249`, and `y10:AA123` are all valid;
  ranges are ordered automatically (`C13:A2` == `A2:C13`).
- **Output order follows option order.** Extractions print in the order given,
  separated by a single blank line. Range columns join with ` -- ` (configurable
  in `config.json`).

### Examples

```bash
# A single cell to the terminal
extract-excel book.xlsx --cell A2

# Several cells (mixed case is fine)
extract-excel book.xlsx --cell C5 -c r249 --cell yc30 -c H7

# A range, and multiple ranges
extract-excel book.xlsx --range A2:F45
extract-excel book.xlsx --range A2:G5 -r B4:U8 --range y10:AA123

# A column by its header title (header row auto-detected, or pinned)
extract-excel book.xlsx --title "Company Name"
extract-excel book.xlsx --title row=2 "March"

# A whole sheet's used range (first data cell -> last data cell)
extract-excel book.xlsx --range sheet                 # single-sheet workbook
extract-excel book.xlsx -r sheet:"Sales Dashboard"    # a named sheet

# Pick a sheet (must immediately follow the workbook)
extract-excel book.xlsx -s "Monthly Budget" --range B5:N9

# Inspect a workbook's sheets (no extraction)
extract-excel book.xlsx --sheet:length             # number of sheets
extract-excel book.xlsx --sheet:list               # one name per line
extract-excel book.xlsx --sheet:info               # count + names combined
extract-excel book.xlsx --sheet:list --file sheets.txt   # write to a file
extract-excel book.xlsx -s:info                    # short flag form

# Output routing
extract-excel book.xlsx --cell C5 --file out.txt
extract-excel book.xlsx --cell Z22 --action:file,stdout out.txt
extract-excel book.xlsx --title "Company Name" -t Phone --action:pdf summary.pdf
extract-excel book.xlsx --cell Z22 --action:file,var var=_name file=out.txt

# Output formats: text (default), table, csv, markdown, xml
extract-excel book.xlsx --range sheet --action:table          # aligned grid
extract-excel book.xlsx --range A1:F20 --file report.csv      # csv inferred
extract-excel book.xlsx -r sheet:"Pivot" --action:markdown    # markdown table

# XML: the sheet mapped to records (header row names the fields)
extract-excel book.xlsx --xml                                 # whole sheet -> terminal
extract-excel book.xlsx -s "Pivot" --xml --file pivot.xml     # to a file
extract-excel book.xlsx -s XML --xml:cardholders,CardHolder   # override root/row tags

# Aligned PDF / markdown tables (extension implies the target)
extract-excel book.xlsx --range sheet --file report.pdf       # pdf target
extract-excel book.xlsx --range sheet --file report.md        # md target
extract-excel book.xlsx --range sheet -o landscape --action:pdf report.pdf
extract-excel book.xlsx --range sheet --assume-merge --file summary.md

# Print the installed version
extract-excel --version

# Many workbooks, many destinations, one command
extract-excel a.xlsx -c A1 --file a.txt \
  -b b.xlsx --range B8:AB25 --cell A1 --action:stdout,file b.txt \
  -b c.xlsx --cell A8 --action:var _value
```

### Output actions

`--action` takes colon-joined **targets** and optional **format** modifiers, and
routes a workbook's combined output to each target. Targets that need a
path/name are `file`, `pdf`, `md`, and `var`.

| Form                                            | Meaning                                  |
| :---------------------------------------------- | :--------------------------------------- |
| `--action:stdout`                               | Terminal (the default).                  |
| `--action:file out.txt`                         | Single target → next token is its value. |
| `--action:file,stdout out.txt`                  | File **and** terminal.                   |
| `--action:table`                                | Terminal, formatted as an aligned table. |
| `--action:xml`                                  | Terminal, formatted as XML.              |
| `--action:xml out.xml`                          | File, formatted as XML.                  |
| `--action:file,csv out.csv`                     | File, formatted as CSV.                  |
| `--action:pdf report.pdf`                       | Aligned table rendered into a PDF.       |
| `--action:md report.md`                         | Aligned table written to a `.md` file.   |
| `--action:file,var var=_n file=out.txt`         | Multiple value targets → `name=value`.   |
| `--action:pdf,file,var var=_n pdf=s.pdf file=o.txt` | Three targets, each named.           |

**Formats** — `text` (default), `table`, `csv`, `markdown`, `xml`. Add one as a
modifier among the targets (e.g. `--action:file,table`). When writing a plain
file, the format is also inferred from the extension (`.csv` → csv, `.md` →
markdown, `.xml` → xml). A format-only action followed by a path
(`--action:xml out.xml`) writes that file; with no path it prints to the
terminal.

#### XML (`-x, --xml`)

`-x/--xml` is shorthand for the `--action:xml` format modifier. It maps a
sheet's table to records: the **header row names the fields** (`Last Name` →
`<Last_Name>`, `FT/PT` → `<FT_PT>`) and every later row becomes a record. Blank
rows are skipped. With no `-c/-r/-t` op, `--xml` extracts the whole sheet (its
used range). On its own it prints to the terminal; pair it with `--file out.xml`
to write a file.

**Container tags** — if the workbook carries an **embedded Excel XML map**
(*Developer → XML*, the same mechanism behind `xmlns:xsi`), `--xml` detects and
reuses its root and repeating-row element names and stamps the `xmlns:xsi`
namespace, so the output matches what Excel itself exports. Otherwise the root is
the camel-cased sheet name (`On Boarding` → `<onBoarding>`) and each record is a
`<row>`. Override either name with `--xml:root,row`
(`--xml:cardholders,CardHolder`, or `--xml:,Record` for just the row).

```xml
<!-- workbook with an embedded XML map -->
<customers xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Customer>
    <Name>Acme Center</Name>
    <Number>-</Number>
    <Status>Active</Status>
  </Customer>
</customers>

<!-- workbook with no map: sheet name + <row> -->
<data>
  <row>
    <ID>1</ID>
    <Last_Name>Anderson</Last_Name>
    <First_Name>James</First_Name>
  </row>
</data>
```

#### Aligned tables (`pdf` and `md`)

The `pdf` and `md` targets share one renderer that builds a monospace,
pipe-bordered table in two passes — first measuring the widest cell in each
column, then padding every cell to that width so the columns line up:

```text
| ID | Last Name | Department | Salary |
|----|-----------|------------|--------|
| 1  | Doe       | Finance    | 62000  |
| 2  | Smith     | IT         | 87000  |
```

A fully empty row splits a sheet into stacked tables, each with its own columns
and widths, and columns that are empty across a whole table are dropped — so a
summary block beneath a blank row renders as its own clean table rather than
being stretched to the grid above it. Only the first table receives a markdown
header rule.

`--file out.pdf` and `--file out.md` route through this renderer automatically;
any other final extension (including `out.md.txt`) stays a plain text file.

#### `--assume-merge`

When a cell is merged across several columns, the spanned value repeats in the
grid (`Total | Total | …`). Pass `--assume-merge` to collapse those repeated
**text** cells back into one; purely numeric duplicates are left alone:

```text
# without --assume-merge
| Headcount | Headcount | 50 | Payroll | Payroll | 4464930 |

# with --assume-merge
| Headcount | 50 | Payroll | 4464930 |
```

It accepts an explicit value (`--assume-merge=true` / `--assume-merge=false`) and
defaults to off. Used once anywhere within a book it applies to every extraction
in that book; used more than once, place it before each extraction to vary the
setting.

#### Sheet queries (`--sheet:length`, `:list`, `:info`)

The qualifier forms of `-s/--sheet` inspect a workbook's sheet structure without
running any extraction. They can be routed to any output target with `-f/--file`
or `--action`.

| Qualifier        | Output                                                        |
| :--------------- | :------------------------------------------------------------ |
| `--sheet:length` | Bare sheet count (e.g. `3`).                                  |
| `--sheet:list`   | One sheet name per line; markdown bullet list for `pdf`/`md`. |
| `--sheet:info`   | `length: N` header followed by `- Name` bullets.             |

```bash
extract-excel book.xlsx --sheet:length
# 3

extract-excel book.xlsx --sheet:list
# Sheet1
# Sheet2
# Sheet3

extract-excel book.xlsx --sheet:info
# length: 3
# - Sheet1
# - Sheet2
# - Sheet3

# Write to a file (the .md extension triggers bullet-list format)
extract-excel book.xlsx --sheet:list --file sheets.md
```

#### `-o, --orientation`

Sets the PDF page layout to `landscape` or `portrait`. It only affects PDF
output (a `pdf` target or a `.pdf` file) and is ignored otherwise:

```bash
extract-excel book.xlsx --range sheet -o landscape --file report.pdf
```

`var` emits an OS-appropriate, sourceable assignment (Windows `set "..."`,
PowerShell `$env:..`, POSIX `export ..`) because a child process cannot set a
variable in its parent shell. Capture it with the bundled helpers:

```bash
# bash — source it so the variable lands in your shell
source scripts/ee-capture.sh book.xlsx --cell B2 --action:var _lastName
echo "$_lastName"
```

```bat
:: cmd — call it
call scripts\ee-capture.cmd book.xlsx --cell B2 --action:var _lastName
echo %_lastName%
```

```powershell
# PowerShell — dot-source it
. .\scripts\ee-capture.ps1 book.xlsx --cell B2 --action:var _lastName
$env:_lastName
```

### The `--range sheet` preset

`--range sheet` extracts a sheet's **used range** — the bounding box from its
first cell with data to its last. Use the bare form for single-sheet workbooks,
or `-r sheet:"Sheet Name"` to target a named sheet in a multi-sheet workbook
(no separate `--sheet` needed):

```bash
extract-excel book.xlsx --range sheet
extract-excel report.xlsx -r sheet:"Monthly Budget" --action:markdown
```

---

## API usage

```ts
import {
  extract,
  loadWorkbook,
  extractRange,
  extractTitle,
} from 'extract-excel';

// High-level: run a CLI-style command and capture the stdout text.
const text = await extract(['book.xlsx', '--range', 'A1:C5']);

// Low-level: drive the engine directly.
const workbook = await loadWorkbook('book.xlsx');
const sheet = workbook.resolveSheet('Sheet1');
const ages = extractTitle(sheet, 'Age', { config: loadConfig() });
const block = extractRange(sheet, 'A2:C13');
```

Key exports: `extract`, `parse`, `runExtract`, `runTests`, `loadWorkbook`,
`workbookFromCsvText`, `extractCell`, `extractRange`, `extractTitle`,
`render`/`renderAligned` and the other renderers, `getVersion`, `Workbook`,
`Sheet`, address helpers (`parseCell`, `parseRange`, …), and the `ExtractError`
type for programmatic error handling.

---

## Testing

The tool ships its own test subsystem behind `--test`:

```bash
extract-excel --test                       # full suite
extract-excel --test unit                  # all unit tests
extract-excel --test unit:extract          # unit tests for extraction
extract-excel --test unit:extract type=multi   # multi-value extractors only
extract-excel --test unit:extract opt=range    # the --range option only

# Custom replays against the bundled fixtures or your own data
extract-excel --test custom:single "--cell A3 -c D54"
extract-excel --test custom:multi "--sheet 'Monthly Budget' --range B5:N9"
extract-excel --test custom:"some file.csv" "--cell C3"
extract-excel --test 'custom:Name,Age\nJane Doe,41' "--cell A2 --title Age"
```

A custom test runs and prints data exactly as if the resolved file (or raw CSV)
had been passed on the command line.

---

## Configuration

`config.json` (shipped with the package) controls formatting:

```json
{
  "columnSeparator": " -- ",
  "blockSeparator": "",
  "headerDetection": "first-nonempty",
  "varExport": "auto"
}
```

| Key               | Purpose                                                        |
| :---------------- | :------------------------------------------------------------- |
| `columnSeparator` | Joins columns within a range row.                              |
| `blockSeparator`  | Text between extract outputs (empty → one blank line).         |
| `headerDetection` | How `--title` finds the header row when `row=` is absent.      |
| `varExport`       | `auto` / `cmd` / `powershell` / `posix` for `--action:var`.    |

---

## License

MIT
