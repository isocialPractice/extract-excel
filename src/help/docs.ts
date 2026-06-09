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
  -a, --action       output:multi    Direct output (stdout, file, pdf, md, var).
  -b, --book         file:single     Workbook to extract data from.
  -c, --cell         extract:single  Extract a single cell, e.g. -c A2.
  -f, --file         output:single   Write output to a file (.pdf/.md route by ext).
  -h, --help         app:doc         Show global or specific help.
  -o, --orientation  switch:string   PDF page orientation (landscape | portrait).
  -r, --range        extract:multi   Extract a range, e.g. -r A2:F45.
  -s, --sheet[:q]    select:single   Select a sheet (must follow its workbook).
                                     :length — sheet count; :list — sheet names;
                                     :info  — count + names combined.
  -t, --title        extract:multi   Extract a column by header title.
  -v, --version      app:doc         Print the installed version.
  -x, --xml          output:single   Output the sheet's mapped table as XML.
      --assume-merge switch:bool     Collapse spanned merge cells in pdf/md output.
      --test         app:test        Run global, unit, or custom tests.

Notes:
  - -s/--sheet is only honored immediately after its workbook.
  - --sheet:length returns the number of sheets in a workbook; --sheet:list
    returns each sheet name on its own line; --sheet:info returns both as a
    "length: N" header followed by a "- Name" bullet list. All three can be
    routed with -f/--file. The pdf/md target uses markdown list format for
    :list and :info.
  - A workbook with >1 sheet and no --sheet raises a custom error.
  - --range sheet extracts a single-sheet workbook's whole used range;
    --range sheet:"Name" does the same for a named sheet (multi-sheet).
  - --action targets: stdout, file, pdf, md, var. Combine with commas
    (--action:file,stdout) and pass multiple paths as name=value
    (--action:file,var var=_name file=out.txt).
  - --action formats: text (default), table, csv, markdown, xml. Add one as a
    modifier (--action:file,table). file output also infers csv/markdown/xml
    from the extension.
  - -x/--xml is shorthand for --action:xml. It maps a sheet's table to XML: the
    header row names the fields (Last Name => <Last_Name>) and each later row is
    a record. With no -c/-r/-t op, --xml extracts the whole sheet (its used
    range), skipping blank rows. Bare --xml prints to the terminal; --xml --file
    out.xml (or --action:xml out.xml) writes a file.
  - Container tags: if the workbook has an embedded Excel XML map (Developer >
    XML), --xml uses its root and repeating-row element names and adds
    xmlns:xsi (e.g. <cardholders ...><CardHolder>...). Otherwise the root is the
    camel-cased sheet name and rows are <row>. Override either with
    --xml:root,row (e.g. --xml:cardholders,CardHolder; --xml:,Record sets only
    the row).
  - pdf and md targets share one two-pass aligned table (columns padded to line
    up). A blank row splits a sheet into stacked tables; empty columns drop out.
  - --file out.pdf and --file out.md imply --action:pdf / --action:md; any other
    extension (including out.md.txt) stays a plain text file.
  - --assume-merge (or =true/=false) collapses repeated text from spanned merges
    in pdf/md output; numeric duplicates are kept. Use once anywhere in a book to
    apply to all of its extractions, or before each one to vary it.
  - -o/--orientation sets the PDF page layout; it is ignored without PDF output.
  - Capture --action:var output with scripts/ee-capture.{sh,cmd,ps1}.`;

const EXAMPLES = `Examples:
  extract-excel book.xlsx --cell A2
  extract-excel book.xlsx --sheet:length
  extract-excel book.xlsx --sheet:list
  extract-excel book.xlsx --sheet:info
  extract-excel book.xlsx --sheet:list --file sheets.txt
  extract-excel book.xlsx --range A2:F45 -r B4:U8
  extract-excel book.xlsx --title "Company Name" -t Phone
  extract-excel book.xlsx -s "Monthly Budget" --range B5:N9
  extract-excel book.xlsx --cell C5 --file out.txt
  extract-excel book.xlsx --cell Z22 --action:file,stdout out.txt
  extract-excel book.xlsx --range sheet --action:table
  extract-excel book.xlsx -r sheet:"Sales Dashboard" --action:markdown
  extract-excel book.xlsx --xml
  extract-excel book.xlsx -s XML --xml:cardholders,CardHolder --file out.xml
  extract-excel book.xlsx --range A1:F20 --file report.csv
  extract-excel book.xlsx --range sheet --action:pdf report.pdf
  extract-excel book.xlsx --range sheet --action:pdf --assume-merge report.pdf
  extract-excel book.xlsx --range sheet -o landscape --file report.pdf
  extract-excel book.xlsx --range sheet --file summary.md
  extract-excel a.xlsx -c A1 -b b.xlsx --range B8:AB25 --action:stdout,file b.txt
  extract-excel --version
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
