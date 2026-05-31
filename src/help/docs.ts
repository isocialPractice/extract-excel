/**
 * Help text rendered for `-h/--help` and its specific topics (`opt`, `quick`).
 * Kept as plain strings so the CLI can print them verbatim with no dependency.
 */
import { HelpCommand } from '../parser/types';

const QUICKSTART = `extract-excel — extract Microsoft Excel data

Usage:
  extract-excel [file|app] [option] [task] [arguments...]

Quickstart:
  [1]      File or App        a workbook path, or an app option (--help/--test)
  [2]      Option             an extraction or output flag
  [3]      Task               the data task implied by the option
  [4-*]    repeat             more [file, option, task] groups or arguments

If the first parameter is not an app option, it is treated as -b/--book.`;

const OPTIONS = `Options:
  -a, --action   output:multi    Direct output (stdout, file, pdf, var).
  -b, --book     file:single     Workbook to extract data from.
  -c, --cell     extract:single  Extract a single cell, e.g. -c A2.
  -f, --file     output:single   Write output to a file.
  -h, --help     app:doc         Show global or specific help.
  -r, --range    extract:multi   Extract a range, e.g. -r A2:F45.
  -s, --sheet    select:single   Select a sheet (must follow its workbook).
  -t, --title    extract:multi   Extract a column by header title.
      --test     app:test        Run global, unit, or custom tests.

Notes:
  - -s/--sheet is only honored immediately after its workbook.
  - A workbook with >1 sheet and no --sheet raises a custom error.
  - --range sheet extracts a single-sheet workbook's whole used range;
    --range sheet:"Name" does the same for a named sheet (multi-sheet).
  - --action targets: stdout, file, pdf, var. Combine with commas
    (--action:file,stdout) and pass multiple paths as name=value
    (--action:file,var var=_name file=out.txt).
  - --action formats: text (default), table, csv, markdown. Add one as a
    modifier (--action:file,table). file output also infers csv/markdown
    from the extension; pdf is always a markdown table (merged cells spanned).
  - Capture --action:var output with scripts/ee-capture.{sh,cmd,ps1}.`;

const EXAMPLES = `Examples:
  extract-excel book.xlsx --cell A2
  extract-excel book.xlsx --range A2:F45 -r B4:U8
  extract-excel book.xlsx --title "Company Name" -t Phone
  extract-excel book.xlsx -s "Monthly Budget" --range B5:N9
  extract-excel book.xlsx --cell C5 --file out.txt
  extract-excel book.xlsx --cell Z22 --action:file,stdout out.txt
  extract-excel book.xlsx --range sheet --action:table
  extract-excel book.xlsx -r sheet:"Sales Dashboard" --action:markdown
  extract-excel book.xlsx --range A1:F20 --file report.csv
  extract-excel book.xlsx --range sheet --action:pdf report.pdf
  extract-excel a.xlsx -c A1 -b b.xlsx --range B8:AB25 --action:stdout,file b.txt
  extract-excel --help opt
  extract-excel --test unit:extract type=multi`;

/** Resolve the help text for a parsed help command. */
export function renderHelp(command: HelpCommand): string {
  switch (command.topic) {
    case 'quick':
      return QUICKSTART;
    case 'opt':
      return OPTIONS;
    case 'global':
    default:
      return [QUICKSTART, '', OPTIONS, '', EXAMPLES].join('\n');
  }
}
