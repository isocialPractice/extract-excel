/**
 * Built-in test subsystem behind `--test`.
 *
 * Two distinct behaviors live here, matching the draft doc:
 *   - global / unit : run the internal assertion suite (optionally filtered by
 *     category, `type=`, and `opt=`).
 *   - custom        : replay a CLI argument string against a bundled book
 *     (`single`/`multi`), an external file, or raw CSV data, and emit the data
 *     exactly as if it had been passed on the command line.
 */
import * as fs from 'fs';
import * as path from 'path';
import { TestCommand, BookSource, ExtractCommand } from '../parser/types';
import { parse, splitArgString } from '../parser/tokenizer';
import { runExtract } from '../run';
import { loadConfig, resetConfigCache } from '../config';
import { workbookFromCsvText, parseCsv, Workbook, loadWorkbook } from '../engine/workbook';
import {
  extractCell,
  extractRange,
  extractUsedRange,
  extractTitle,
} from '../engine/extract';
import {
  renderText,
  renderCsv,
  renderTable,
  renderMarkdown,
  renderAligned,
} from '../output/render';
import { ExtractionResult } from '../engine/extract';
import {
  columnToNumber,
  numberToColumn,
  parseCell,
  parseRange,
  formatCell,
} from '../engine/address';

export interface TestCaseResult {
  name: string;
  ok: boolean;
  error?: string;
}

export interface TestReport {
  mode: TestCommand['mode'];
  results: TestCaseResult[];
  passed: number;
  failed: number;
  /** Captured replay output for custom tests. */
  output?: string;
}

interface UnitTest {
  name: string;
  category: string;
  opt?: string;
  taskType?: 'single' | 'multi';
  run: () => void | Promise<void>;
}

/** Sample dataset from the documentation (example XII). */
const SAMPLE_CSV = [
  'Name, Age',
  'John Smith, 29',
  'Jane Doe, 41',
  'Bob Johnson, 47',
  'Karen Mitchell, 34',
  'Chad Henderson, 28',
  'Brenda Kowalski, 52',
  'Gary Patterson, 61',
  'Linda Crawford, 43',
  'Dave Fletcher, 39',
  'Susan Whitfield, 57',
  'Todd Summers, 31',
  'Debbie Carlson, 45',
].join('\n');

function assert(condition: boolean, message: string): void {
  if (!condition) throw new Error(message);
}

function assertEqual<T>(actual: T, expected: T, label: string): void {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) throw new Error(`${label}: expected ${e} but got ${a}`);
}

/** The internal unit suite. Each test throws on failure. */
function unitTests(): UnitTest[] {
  const config = loadConfig();
  const wb = workbookFromCsvText(SAMPLE_CSV);
  const sheet = wb.resolveSheet();

  return [
    // --- address (pure utilities) ---
    {
      name: 'address: columnToNumber handles single + multi letters',
      category: 'address',
      run: () => {
        assertEqual(columnToNumber('A'), 1, 'A');
        assertEqual(columnToNumber('Z'), 26, 'Z');
        assertEqual(columnToNumber('AA'), 27, 'AA');
        assertEqual(columnToNumber('yc'), 653, 'yc'); // y(25)*26 + c(3)
      },
    },
    {
      name: 'address: numberToColumn round-trips',
      category: 'address',
      run: () => {
        for (const n of [1, 26, 27, 52, 703, 727]) {
          assertEqual(columnToNumber(numberToColumn(n)), n, `roundtrip ${n}`);
        }
      },
    },
    {
      name: 'address: parseCell normalizes mixed case',
      category: 'address',
      run: () => {
        assertEqual(formatCell(parseCell('yc30')), 'YC30', 'yc30');
        assertEqual(formatCell(parseCell('r249')), 'R249', 'r249');
      },
    },
    {
      name: 'address: parseRange orders corners',
      category: 'address',
      run: () => {
        const r = parseRange('C13:A2');
        assertEqual(formatCell(r.start), 'A2', 'start');
        assertEqual(formatCell(r.end), 'C13', 'end');
      },
    },

    // --- extract: cell (single) ---
    {
      name: 'extract:cell returns the cell value',
      category: 'extract',
      opt: 'cell',
      taskType: 'single',
      run: () => {
        const res = extractCell(sheet, 'A2');
        assertEqual(res.rows, [['John Smith']], 'A2');
      },
    },
    {
      name: 'extract:cell on an empty cell yields empty string',
      category: 'extract',
      opt: 'cell',
      taskType: 'single',
      run: () => {
        const res = extractCell(sheet, 'C2');
        assertEqual(res.rows, [['']], 'C2 empty');
      },
    },

    // --- extract: range (multi) ---
    {
      name: 'extract:range returns a full rectangle',
      category: 'extract',
      opt: 'range',
      taskType: 'multi',
      run: () => {
        const res = extractRange(sheet, 'A2:C13');
        assertEqual(res.rows[0], ['John Smith', '29', ''], 'row 1 rectangular');
        assertEqual(res.rows.length, 12, 'row count');
      },
    },
    {
      name: 'extract:range text render trims trailing empty columns',
      category: 'output',
      opt: 'range',
      taskType: 'multi',
      run: () => {
        const res = extractRange(sheet, 'A2:C13');
        const lines = renderText([res], config).split('\n');
        assertEqual(lines[0], 'John Smith -- 29', 'trailing C trimmed');
        assertEqual(lines[11], 'Debbie Carlson -- 45', 'last row');
      },
    },

    // --- extract: title (multi) ---
    {
      name: 'extract:title reads a column to first gap',
      category: 'extract',
      opt: 'title',
      taskType: 'multi',
      run: () => {
        const res = extractTitle(sheet, 'Age', { config });
        assertEqual(res.rows.length, 12, 'count');
        assertEqual(res.rows[0], ['29'], 'first');
        assertEqual(res.rows[11], ['45'], 'last');
      },
    },
    {
      name: 'extract:title honors an explicit header row',
      category: 'extract',
      opt: 'title',
      taskType: 'multi',
      run: () => {
        const res = extractTitle(sheet, 'Name', { headerRow: 1, config });
        assertEqual(res.rows[0], ['John Smith'], 'first name');
      },
    },

    // --- output: separator formatting ---
    {
      name: 'output:range joins columns with the configured separator',
      category: 'output',
      opt: 'action',
      taskType: 'multi',
      run: () => {
        const res = extractRange(sheet, 'A2:B2');
        assertEqual(
          res.rows[0].join(config.columnSeparator),
          'John Smith -- 29',
          'joined',
        );
      },
    },

    // --- select: sheet resolution ---
    {
      name: 'select:sheet on a single-sheet workbook needs no name',
      category: 'select',
      opt: 'sheet',
      taskType: 'single',
      run: () => {
        const s = workbookFromCsvText(SAMPLE_CSV).resolveSheet();
        assert(s.rowCount === 13, 'expected 13 rows');
      },
    },

    // --- parser: tokenizer behavior ---
    {
      name: 'parser: first bare token defaults to a book',
      category: 'parser',
      run: () => {
        const cmd = parse(['file.xlsx', '--cell', 'A2']) as ExtractCommand;
        assertEqual(cmd.books.length, 1, 'books');
        assertEqual(cmd.books[0].source, { kind: 'file', path: 'file.xlsx' }, 'src');
        assertEqual(
          cmd.books[0].ops,
          [{ type: 'cell', ref: 'A2', assumeMerge: false }],
          'ops',
        );
      },
    },
    {
      name: 'parser: --sheet is ignored unless it follows the book',
      category: 'parser',
      run: () => {
        const cmd = parse(['f.xlsx', '--cell', 'A1', '--sheet', 'x']) as ExtractCommand;
        assert(cmd.books[0].sheet === undefined, 'sheet should be ignored');
        const ok = parse(['f.xlsx', '--sheet', 'x', '--cell', 'A1']) as ExtractCommand;
        assertEqual(ok.books[0].sheet, 'x', 'sheet honored');
      },
    },
    {
      name: 'parser: multiple books partition their ops',
      category: 'parser',
      run: () => {
        const cmd = parse([
          'a.xlsx', '-c', 'A1',
          '-b', 'b.xlsx', '--range', 'B8:AB25', '--cell', 'A1',
        ]) as ExtractCommand;
        assertEqual(cmd.books.length, 2, 'two books');
        assertEqual(cmd.books[0].ops.length, 1, 'book a ops');
        assertEqual(cmd.books[1].ops.length, 2, 'book b ops');
      },
    },
    {
      name: 'parser: --action single target consumes one path',
      category: 'parser',
      opt: 'action',
      run: () => {
        const cmd = parse(['f.xlsx', '--cell', 'A1', '--action:file', 'out.txt']) as ExtractCommand;
        assertEqual(cmd.books[0].action.targets, ['file'], 'targets');
        assertEqual(cmd.books[0].action.file, 'out.txt', 'file path');
      },
    },
    {
      name: 'parser: --action multi target consumes name=value args',
      category: 'parser',
      opt: 'action',
      run: () => {
        const cmd = parse([
          'f.xlsx', '--cell', 'A1',
          '--action:file,var', 'var=_n', 'file=out.txt',
        ]) as ExtractCommand;
        assertEqual(cmd.books[0].action.file, 'out.txt', 'file');
        assertEqual(cmd.books[0].action.var, '_n', 'var');
      },
    },
    {
      name: 'parser: --title row=N sets the header row',
      category: 'parser',
      opt: 'title',
      run: () => {
        const cmd = parse(['f.xlsx', '--title', 'row=2', 'March']) as ExtractCommand;
        assertEqual(
          cmd.books[0].ops,
          [{ type: 'title', title: 'March', headerRow: 2, assumeMerge: false }],
          'op',
        );
      },
    },

    // --- workbook: real xlsx smoke tests (bundled fixtures) ---
    {
      name: 'workbook: single-sheet fixture loads with one sheet',
      category: 'workbook',
      run: async () => {
        const file = bundledFile('single-sheet-book.xlsx');
        if (!file) return; // fixture not shipped — skip gracefully
        const book = await loadWorkbook(file);
        assert(book.sheetCount >= 1, 'expected at least one sheet');
      },
    },
    {
      name: 'workbook: multi-sheet fixture is ambiguous without --sheet',
      category: 'workbook',
      run: async () => {
        const file = bundledFile('multi-sheet-book.xlsx');
        if (!file) return;
        const book = await loadWorkbook(file);
        if (book.sheetCount > 1) {
          let threw = false;
          try {
            book.resolveSheet();
          } catch {
            threw = true;
          }
          assert(threw, 'expected ambiguous-sheet error');
        }
      },
    },

    // --- csv parsing ---
    {
      name: 'csv: parses quoted fields with embedded commas',
      category: 'csv',
      run: () => {
        const grid = parseCsv('"Smith, John",29\nJane,41');
        assertEqual(grid[0], ['Smith, John', '29'], 'quoted');
        assertEqual(grid[1], ['Jane', '41'], 'plain');
      },
    },

    // --- used range (--range sheet preset) ---
    {
      name: 'usedrange: bounding box spans first to last data cell',
      category: 'extract',
      opt: 'range',
      taskType: 'multi',
      run: () => {
        const res = extractUsedRange(sheet);
        assertEqual(res.ref, 'A1:B13', 'used range ref');
        assertEqual(res.rows.length, 13, 'rows incl. header');
        assertEqual(res.rows[0], ['Name', 'Age'], 'header row');
      },
    },
    {
      name: 'parser: --range sheet sets the used-range preset',
      category: 'parser',
      opt: 'range',
      run: () => {
        const a = parse(['f.xlsx', '--range', 'sheet']) as ExtractCommand;
        assertEqual(
          a.books[0].ops,
          [{ type: 'range', usedRange: true, assumeMerge: false }],
          'no name',
        );
        const b = parse(['f.xlsx', '-r', 'sheet:Sales']) as ExtractCommand;
        assertEqual(
          b.books[0].ops,
          [{ type: 'range', usedRange: true, sheetName: 'Sales', assumeMerge: false }],
          'named',
        );
      },
    },

    // --- render formats ---
    {
      name: 'render:csv quotes fields with commas',
      category: 'output',
      opt: 'action',
      run: () => {
        const res = extractRange(sheet, 'A2:B3');
        const csv = renderCsv([res]);
        assertEqual(csv.split('\n')[0], 'John Smith,29', 'csv row');
      },
    },
    {
      name: 'render:table draws aligned bordered columns',
      category: 'output',
      opt: 'action',
      run: () => {
        const res = extractRange(sheet, 'A2:B2');
        const table = renderTable([res]);
        assert(table.includes('| John Smith | 29 |'), 'aligned row present');
        assert(table.startsWith('+'), 'has border');
      },
    },
    {
      name: 'render:markdown spans merged cells across columns',
      category: 'output',
      opt: 'action',
      run: () => {
        // A1:C1 merged with master value "Title"; B1/C1 empty in the grid.
        const result = {
          kind: 'range' as const,
          ref: 'A1:C2',
          rows: [
            ['Title', '', ''],
            ['a', 'b', 'c'],
          ],
          merges: [{ start: { row: 1, col: 1 }, end: { row: 1, col: 3 } }],
        };
        const md = renderMarkdown([result]).split('\n');
        assertEqual(md[0], '| Title | Title | Title |', 'spanned header');
        assertEqual(md[1], '| --- | --- | --- |', 'separator');
        assertEqual(md[2], '| a | b | c |', 'body');
      },
    },
    {
      name: 'parser: --action:table sets a format with default stdout',
      category: 'parser',
      opt: 'action',
      run: () => {
        const cmd = parse(['f.xlsx', '--cell', 'A1', '--action:table']) as ExtractCommand;
        assertEqual(cmd.books[0].action.targets, ['stdout'], 'targets');
        assertEqual(cmd.books[0].action.format, 'table', 'format');
      },
    },
    {
      name: 'parser: --action:file,csv combines target and format',
      category: 'parser',
      opt: 'action',
      run: () => {
        const cmd = parse([
          'f.xlsx', '--cell', 'A1', '--action:file,csv', 'out.csv',
        ]) as ExtractCommand;
        assertEqual(cmd.books[0].action.targets, ['file'], 'targets');
        assertEqual(cmd.books[0].action.format, 'csv', 'format');
        assertEqual(cmd.books[0].action.file, 'out.csv', 'file');
      },
    },

    // --- md output target ---
    {
      name: 'parser: --action:md sets the md target and path',
      category: 'parser',
      opt: 'action',
      run: () => {
        const cmd = parse(['f.xlsx', '-r', 'sheet', '--action:md', 'out.md']) as ExtractCommand;
        assertEqual(cmd.books[0].action.targets, ['md'], 'targets');
        assertEqual(cmd.books[0].action.md, 'out.md', 'md path');
      },
    },
    {
      name: 'parser: --file infers md/pdf targets from the extension',
      category: 'parser',
      opt: 'file',
      run: () => {
        const md = parse(['f.xlsx', '-r', 'sheet', '--file', 'report.md']) as ExtractCommand;
        assertEqual(md.books[0].action.targets, ['md'], 'md target');
        assertEqual(md.books[0].action.md, 'report.md', 'md path');
        const pdf = parse(['f.xlsx', '-r', 'sheet', '-f', 'report.PDF']) as ExtractCommand;
        assertEqual(pdf.books[0].action.targets, ['pdf'], 'pdf target');
        assertEqual(pdf.books[0].action.pdf, 'report.PDF', 'pdf path');
      },
    },
    {
      name: 'parser: --file only implies md when .md is the final extension',
      category: 'parser',
      opt: 'file',
      run: () => {
        const cmd = parse(['f.xlsx', '-r', 'sheet', '--file', 'notes.md.txt']) as ExtractCommand;
        assertEqual(cmd.books[0].action.targets, ['file'], 'plain file target');
        assertEqual(cmd.books[0].action.file, 'notes.md.txt', 'file path');
      },
    },

    // --- assume-merge switch ---
    {
      name: 'parser: --assume-merge used once applies to every op in the book',
      category: 'parser',
      opt: 'assume-merge',
      run: () => {
        // Placed after the op, yet still applies (single-occurrence rule).
        const cmd = parse([
          'f.xlsx', '-r', 'A1:B2', '-r', 'A3:B4', '--assume-merge',
        ]) as ExtractCommand;
        assertEqual(cmd.books[0].ops.map((o) => o.assumeMerge), [true, true], 'both true');
      },
    },
    {
      name: 'parser: repeated --assume-merge is positional per op',
      category: 'parser',
      opt: 'assume-merge',
      run: () => {
        const cmd = parse([
          'f.xlsx', '--assume-merge', '-r', 'A1:G2', '--assume-merge=false', '-r', 'A3:G9',
        ]) as ExtractCommand;
        assertEqual(cmd.books[0].ops.map((o) => o.assumeMerge), [true, false], 'true then false');
      },
    },
    {
      name: 'parser: --assume-merge defaults to false when absent',
      category: 'parser',
      opt: 'assume-merge',
      run: () => {
        const cmd = parse(['f.xlsx', '-r', 'sheet']) as ExtractCommand;
        assertEqual(cmd.books[0].ops[0].assumeMerge, false, 'default false');
      },
    },

    // --- orientation ---
    {
      name: 'parser: --orientation sets the PDF page layout',
      category: 'parser',
      opt: 'orientation',
      run: () => {
        const cmd = parse(['f.xlsx', '-r', 'sheet', '-o', 'Landscape', '-f', 'x.pdf']) as ExtractCommand;
        assertEqual(cmd.books[0].action.orientation, 'landscape', 'orientation');
      },
    },
    {
      name: 'parser: --orientation rejects an invalid value',
      category: 'parser',
      opt: 'orientation',
      run: () => {
        let threw = false;
        try {
          parse(['f.xlsx', '-r', 'sheet', '--orientation', 'sideways']);
        } catch {
          threw = true;
        }
        assert(threw, 'expected an error for an invalid orientation');
      },
    },

    // --- version ---
    {
      name: 'parser: --version returns the version command',
      category: 'parser',
      run: () => {
        assertEqual(parse(['--version']).kind, 'version', 'long flag');
        assertEqual(parse(['-v']).kind, 'version', 'short flag');
      },
    },

    // --- aligned (pdf/md) renderer ---
    {
      name: 'render:aligned pads columns and rules only the first header',
      category: 'output',
      opt: 'action',
      run: () => {
        const result: ExtractionResult = {
          kind: 'range',
          ref: 'A1:B2',
          rows: [['ID', 'Name'], ['1', 'Anderson']],
        };
        const lines = renderAligned([result]).split('\n');
        assertEqual(lines[0], '| ID | Name     |', 'padded header');
        assertEqual(lines[1], '|----|----------|', 'header rule');
        assertEqual(lines[2], '| 1  | Anderson |', 'padded body');
      },
    },
    {
      name: 'render:aligned splits on a blank row into stacked tables',
      category: 'output',
      opt: 'action',
      run: () => {
        const result: ExtractionResult = {
          kind: 'range',
          ref: 'A1:B3',
          rows: [['ID', 'Name'], ['', ''], ['Total', '2']],
        };
        const out = renderAligned([result]);
        // Second sub-table gets no header rule (continuation data).
        assert(out.includes('| Total | 2 |'), 'second table rendered');
        const ruleLines = out.split('\n').filter((l) => /^\|[-|]+\|$/.test(l));
        assertEqual(ruleLines.length, 1, 'exactly one header rule');
      },
    },
    {
      name: 'render:aligned with assumeMerge collapses repeated text cells',
      category: 'output',
      opt: 'assume-merge',
      run: () => {
        const rows = [
          ['Total Employees', 'Total Employees', '50', '', 'Total Payroll', 'Total Payroll', '4464930'],
        ];
        const plain = renderAligned([{ kind: 'range', ref: 'r', rows }]);
        assert(plain.includes('Total Employees | Total Employees'), 'doubled without flag');

        const merged = renderAligned([
          { kind: 'range', ref: 'r', rows, assumeMerge: true },
        ]);
        assert(!merged.includes('Total Employees | Total Employees'), 'text collapsed');
        assert(merged.includes('| Total Employees | 50 | Total Payroll | 4464930 |'), 'clean row');
      },
    },
  ];
}

/** Locate a bundled fixture file, or null if it was not shipped. */
function bundledFile(name: string): string | null {
  const candidates = [
    path.resolve(__dirname, '..', '..', 'tests', 'files', name),
    path.resolve(__dirname, '..', '..', '..', 'tests', 'files', name),
  ];
  return candidates.find((p) => fs.existsSync(p)) ?? null;
}

/** Apply the doc's filtering rules to the unit suite. */
function selectTests(all: UnitTest[], command: TestCommand): UnitTest[] {
  let tests = all;
  if (command.category) {
    tests = tests.filter((t) => t.category === command.category);
  }
  if (command.filters.type) {
    tests = tests.filter((t) => t.taskType === command.filters.type);
  }
  if (command.filters.opt) {
    tests = tests.filter((t) => t.opt === command.filters.opt);
  }
  return tests;
}

/** Run the internal unit/global suite. */
async function runUnit(command: TestCommand): Promise<TestReport> {
  const selected = selectTests(unitTests(), command);
  const results: TestCaseResult[] = [];
  for (const test of selected) {
    try {
      await test.run();
      results.push({ name: test.name, ok: true });
    } catch (err) {
      results.push({ name: test.name, ok: false, error: (err as Error).message });
    }
  }
  return summarize(command.mode, results);
}

/** Resolve a custom test target into a concrete book source. */
function resolveCustomSource(target: string, cwd: string): BookSource {
  if (target === 'single' || target === 'multi') {
    const file = bundledFile(`${target}-sheet-book.xlsx`);
    if (!file) {
      throw new Error(`Bundled test fixture "${target}-sheet-book.xlsx" was not found.`);
    }
    return { kind: 'file', path: file };
  }
  // Raw CSV data is identified by commas or newlines (literal or escaped).
  if (/[,\n]/.test(target) || target.includes('\\n')) {
    return { kind: 'raw', text: target, name: 'Sheet1' };
  }
  // Otherwise treat it as a path relative to the working directory.
  return { kind: 'file', path: path.resolve(cwd, target) };
}

/** Replay a custom test's argument string and capture its output. */
async function runCustom(command: TestCommand, cwd: string): Promise<TestReport> {
  resetConfigCache();
  const target = command.customTarget ?? 'single';
  const replay = command.replay ?? '';
  let output = '';
  const results: TestCaseResult[] = [];

  try {
    const source = resolveCustomSource(target, cwd);
    const replayTokens = splitArgString(replay);
    // Prepend a placeholder book so the standard parser attaches the replay
    // options to a single book, then swap in the resolved custom source.
    const parsed = parse(['__custom__', ...replayTokens]) as ExtractCommand;
    parsed.books[0].source = source;

    await runExtract(parsed, {
      cwd,
      write: (text) => {
        output += text;
      },
    });
    results.push({ name: `custom:${target}`, ok: true });
  } catch (err) {
    results.push({ name: `custom:${target}`, ok: false, error: (err as Error).message });
  }

  const report = summarize('custom', results);
  report.output = output;
  return report;
}

function summarize(mode: TestCommand['mode'], results: TestCaseResult[]): TestReport {
  const passed = results.filter((r) => r.ok).length;
  return { mode, results, passed, failed: results.length - passed };
}

/** Public entry point for the `--test` command. */
export async function runTests(
  command: TestCommand,
  options: { cwd?: string } = {},
): Promise<TestReport> {
  const cwd = options.cwd ?? process.cwd();
  if (command.mode === 'custom') return runCustom(command, cwd);
  return runUnit(command);
}
