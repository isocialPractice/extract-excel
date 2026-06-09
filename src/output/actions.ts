/**
 * Output dispatch: render a book's results and send them to each target.
 *
 * Format resolution per target:
 *   - pdf    : the source workbook is exported directly to PDF via LibreOffice
 *              headless (`libreoffice --headless --convert-to pdf`). Requires
 *              LibreOffice to be installed with `libreoffice` or `soffice` on
 *              the PATH. Raw CSV sources are not supported for this target.
 *              When a sheet is selected (`-s/--sheet`), an orientation is set
 *              (`-o/--orientation`), or `--fit` is requested, the workbook is
 *              first rewritten with only the selected sheet and the requested
 *              page setup, then that copy is exported — so the PDF honors the
 *              sheet selection and page layout instead of defaulting to the
 *              workbook's first sheet.
 *   - md     : the two-pass aligned table (same layout the table renderer uses
 *              for the terminal, padded to line up) written to a `.md` file.
 *              `--assume-merge` collapses spanned-merge text in this output.
 *   - file   : the action's explicit format, else inferred from the extension
 *              (`.csv` => csv, `.md` => markdown), else text.
 *   - stdout : the action's explicit format, else text.
 *   - var    : always text, collapsed to a single sourceable assignment. A
 *              child process cannot set a variable in its parent shell, so the
 *              snippet is emitted for the user to `eval` / `for /f` (see the
 *              capture helpers under `scripts/`).
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { ActionSpec, PageOrientation } from '../parser/types';
import { ExtractConfig } from '../config';
import { ExtractError } from '../errors';
import { ExtractionResult } from '../engine/extract';
import { render, renderAligned, formatFromPath } from './render';

const execAsync = promisify(exec);

export interface OutputContext {
  config: ExtractConfig;
  /** Sink for stdout (overridable so the test runner can capture output). */
  write: (text: string) => void;
  /**
   * Resolved absolute path of the source workbook file. Populated only for
   * file-based sources; `undefined` for raw CSV input. Required for the `pdf`
   * target (LibreOffice export).
   */
  sourcePath?: string;
  /**
   * Selected sheet name for this book (`-s/--sheet`), if any. Used by the pdf
   * target to export the requested sheet rather than the workbook's first.
   */
  sheet?: string;
}

/**
 * Dispatch a pre-rendered text string to all configured targets. Used by sheet
 * queries (`--sheet:length`, `:list`, `:info`) which produce plain text, not
 * tabular extraction results. The `pdf` target is silently skipped because a
 * LibreOffice export is not applicable for metadata-only output.
 */
export async function dispatchText(
  text: string,
  action: ActionSpec,
  ctx: OutputContext,
): Promise<void> {
  const targets = action.targets.length > 0 ? action.targets : ['stdout'];
  for (const target of targets) {
    switch (target) {
      case 'stdout':
        ctx.write(text + '\n');
        break;
      case 'file':
        writeTextFile(requirePath(action.file, 'file'), text);
        break;
      case 'md':
        writeTextFile(requirePath(action.md, 'md'), text);
        break;
      case 'var':
        ctx.write(renderVarExport(requireName(action.var), text, ctx.config) + '\n');
        break;
      case 'pdf':
        // Sheet query output is plain text; the pdf target is not applicable.
        break;
    }
  }
}

/** Dispatch one book's results to all of its targets. */
export async function dispatch(
  results: ExtractionResult[],
  action: ActionSpec,
  ctx: OutputContext,
): Promise<void> {
  const targets = action.targets.length > 0 ? action.targets : ['stdout'];
  for (const target of targets) {
    switch (target) {
      case 'stdout':
        ctx.write(render(results, action.format ?? 'text', ctx.config) + '\n');
        break;
      case 'file': {
        const file = requirePath(action.file, 'file');
        const format = action.format ?? formatFromPath(file) ?? 'text';
        writeTextFile(file, render(results, format, ctx.config));
        break;
      }
      case 'pdf': {
        if (!ctx.sourcePath) {
          throw new ExtractError(
            'MALFORMED_ACTION',
            'The pdf target requires a file-based workbook source; raw CSV input cannot be exported as PDF.',
          );
        }
        await exportPdf(ctx.sourcePath, requirePath(action.pdf, 'pdf'), {
          sheet: ctx.sheet,
          orientation: action.orientation,
          fit: action.fit,
        });
        break;
      }
      case 'md':
        // The `.md` target writes the aligned table as plain markdown text.
        writeTextFile(requirePath(action.md, 'md'), renderAligned(results));
        break;
      case 'var': {
        const text = render(results, 'text', ctx.config);
        ctx.write(renderVarExport(requireName(action.var), text, ctx.config) + '\n');
        break;
      }
    }
  }
}

function requirePath(value: string | undefined, target: string): string {
  if (!value) {
    throw new ExtractError('MALFORMED_ACTION', `The "${target}" action has no path.`);
  }
  return value;
}

function requireName(value: string | undefined): string {
  if (!value) {
    throw new ExtractError('MALFORMED_ACTION', 'The "var" action has no variable name.');
  }
  return value;
}

/** Ensure the parent directory exists, then write the text file. */
function writeTextFile(filePath: string, text: string): void {
  const dir = path.dirname(path.resolve(filePath));
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(filePath, text + '\n', 'utf8');
}

/** PDF export options resolved from the action (and book sheet selection). */
interface PdfOptions {
  /**
   * Sheet to isolate and export on its own. When omitted the whole workbook is
   * exported to a single PDF with per-sheet fit + auto-orientation instead.
   */
  sheet?: string;
  /**
   * Explicit page orientation. Applied to the isolated sheet, or to every sheet
   * of a whole-workbook export (overriding the per-sheet auto choice).
   */
  orientation?: PageOrientation;
  /**
   * Scale each exported sheet to a single PDF page. Defaults on for the
   * whole-workbook export; `--fit=false` opts out.
   */
  fit?: boolean;
}

/**
 * Export a workbook to PDF via LibreOffice. Both shapes embed page setup into a
 * temporary copy that LibreOffice then renders:
 *
 *   - With a selected sheet (`-s/--sheet`) only that sheet is exported, carrying
 *     the explicit `-o/--orientation` and `--fit` that were requested.
 *   - With no sheet selected the whole workbook is exported to a single PDF and
 *     every sheet is fitted to one page and auto-oriented from its own used
 *     extent (portrait when taller than wide, landscape when wider, portrait on
 *     a tie). An explicit `-o/--orientation` overrides the auto choice for all
 *     sheets; `--fit=false` turns the per-sheet fit off.
 */
async function exportPdf(
  sourcePath: string,
  outputPath: string,
  opts: PdfOptions,
): Promise<void> {
  const prepared =
    opts.sheet !== undefined
      ? await buildPreparedWorkbook(sourcePath, opts)
      : await buildWholeWorkbookExport(sourcePath, opts);
  try {
    await convertWithLibreOffice(prepared, outputPath);
  } finally {
    fs.rmSync(prepared, { force: true });
  }
}

/**
 * Rewrite `sourcePath` to a temporary `.xlsx` that contains only the selected
 * sheet, with the requested page setup applied. LibreOffice honors the embedded
 * page setup (orientation, fit-to-page) when it converts to PDF.
 */
async function buildPreparedWorkbook(
  sourcePath: string,
  opts: PdfOptions,
): Promise<string> {
  const wb = new ExcelJS.Workbook();
  // Load from a styles-sanitized copy so font flags that are explicitly turned
  // *off* (`<b val="false"/>`, `<i val="0"/>`, `<strike val="false"/>`, …) are
  // dropped before exceljs misreads them as on (see sanitizeFontFlags). Without
  // this the exported PDF would show bold/italic/strikethrough that the cells
  // never had.
  //
  // exceljs types load()'s parameter as ArrayBuffer but accepts a Node Buffer at
  // runtime; cast through unknown to bridge the overly narrow declaration.
  const bytes = (await sanitizedWorkbookBuffer(sourcePath)) as unknown as ArrayBuffer;
  await wb.xlsx.load(bytes);

  const target = resolveWorksheet(wb, sourcePath, opts.sheet);

  // Collapse formulas to their cached values *before* dropping the other sheets.
  // Once a referenced sheet is gone, LibreOffice would recompute a cross-sheet
  // formula to #REF!/#NAME? on load; freezing the last-computed value keeps the
  // number and guarantees no error text reaches the PDF (see flattenFormulas).
  flattenFormulas(target);

  // Drop every other sheet so LibreOffice exports only the selected one.
  for (const ws of [...wb.worksheets]) {
    if (ws.id !== target.id) wb.removeWorksheet(ws.id);
  }

  // Merge the requested page setup over whatever the sheet already defines.
  target.pageSetup = {
    ...target.pageSetup,
    ...(opts.orientation ? { orientation: opts.orientation } : {}),
    ...(opts.fit ? { fitToPage: true, fitToWidth: 1, fitToHeight: 1 } : {}),
  };

  const tmp = tempXlsxPath();
  await wb.xlsx.writeFile(tmp);
  return tmp;
}

/**
 * Rewrite `sourcePath` to a temporary `.xlsx` that keeps *every* sheet, with
 * per-sheet page setup applied so LibreOffice exports the whole workbook to one
 * PDF. Each sheet is fitted to a single page (unless `--fit=false`) and oriented
 * by the explicit `-o/--orientation` or, when none was given, auto from its own
 * used extent (see autoOrientation). Formulas are frozen exactly as in the
 * single-sheet export, so no error text reaches the PDF.
 */
async function buildWholeWorkbookExport(
  sourcePath: string,
  opts: PdfOptions,
): Promise<string> {
  const wb = new ExcelJS.Workbook();
  // Same styles sanitation as the single-sheet path (see sanitizeFontFlags).
  const bytes = (await sanitizedWorkbookBuffer(sourcePath)) as unknown as ArrayBuffer;
  await wb.xlsx.load(bytes);

  if (wb.worksheets.length === 0) {
    throw new ExtractError('UNREADABLE_WORKBOOK', `"${sourcePath}" has no sheets.`);
  }

  // Fit each sheet to a page by default; `--fit=false` opts out (undefined => on).
  const fit = opts.fit ?? true;
  for (const ws of wb.worksheets) {
    flattenFormulas(ws);
    ws.pageSetup = {
      ...ws.pageSetup,
      orientation: opts.orientation ?? autoOrientation(ws),
      ...(fit ? { fitToPage: true, fitToWidth: 1, fitToHeight: 1 } : {}),
    };
  }

  const tmp = tempXlsxPath();
  await wb.xlsx.writeFile(tmp);
  return tmp;
}

/** A unique temp path for the rewritten workbook handed to LibreOffice. */
function tempXlsxPath(): string {
  return path.join(
    os.tmpdir(),
    `extract-excel-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.xlsx`,
  );
}

/** Excel's default column width (character units) and row height (points). */
const DEFAULT_COL_WIDTH = 8.43;
const DEFAULT_ROW_HEIGHT = 15;

/**
 * Choose a page orientation from a sheet's used extent: `landscape` when the
 * data box is physically wider than it is tall, otherwise `portrait` (so a
 * taller-than-wide sheet, a square sheet, and an empty sheet all fall to
 * portrait, matching the requested rule). Column widths (character units) and
 * row heights (points) are converted to a shared pixel scale so the comparison
 * reflects how the sheet would actually print.
 */
export function autoOrientation(ws: ExcelJS.Worksheet): PageOrientation {
  const { widthPx, heightPx } = usedExtentPx(ws);
  return widthPx > heightPx ? 'landscape' : 'portrait';
}

/**
 * Approximate the pixel width and height of a sheet's used cell box. A column's
 * width is in Excel "character" units (~7px per char plus 5px of cell padding
 * for the default font); a row's height is in points (1pt = 4/3 px at 96 DPI).
 * Missing sizes fall back to Excel's defaults. Returns zeros for an empty sheet.
 */
function usedExtentPx(ws: ExcelJS.Worksheet): { widthPx: number; heightPx: number } {
  const defaultColWidth = ws.properties?.defaultColWidth ?? DEFAULT_COL_WIDTH;
  const defaultRowHeight = ws.properties?.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;

  let widthPx = 0;
  for (let c = 1; c <= ws.columnCount; c++) {
    widthPx += Math.round((ws.getColumn(c).width ?? defaultColWidth) * 7 + 5);
  }
  let heightPx = 0;
  for (let r = 1; r <= ws.rowCount; r++) {
    heightPx += (ws.getRow(r).height ?? defaultRowHeight) * (4 / 3);
  }
  return { widthPx, heightPx };
}

/** Resolve the worksheet to export, honoring `--sheet` (case-insensitive). */
function resolveWorksheet(
  wb: ExcelJS.Workbook,
  sourcePath: string,
  sheet?: string,
): ExcelJS.Worksheet {
  if (sheet !== undefined) {
    const found = wb.worksheets.find(
      (ws) => ws.name.toLowerCase() === sheet.toLowerCase(),
    );
    if (!found) {
      throw new ExtractError(
        'SHEET_NOT_FOUND',
        `Sheet "${sheet}" not found in "${sourcePath}".`,
        `Available sheets: ${wb.worksheets.map((ws) => `"${ws.name}"`).join(', ')}`,
      );
    }
    return found;
  }
  const first = wb.worksheets[0];
  if (!first) {
    throw new ExtractError('UNREADABLE_WORKBOOK', `"${sourcePath}" has no sheets.`);
  }
  return first;
}

/**
 * Replace every formula cell on `ws` with its last-computed value, and blank
 * out any cell whose value is a spreadsheet error (`#NAME?`, `#REF!`, `#DIV/0!`,
 * …) or has no result.
 *
 * Two things make this necessary for the PDF export:
 *   - We isolate a single sheet, so a formula referencing a dropped sheet would
 *     otherwise recompute to `#REF!`/`#NAME?`. Freezing the cached result keeps
 *     the value the workbook last calculated.
 *   - Cells with no value (or an error result) should render empty rather than
 *     printing the raw error text in the PDF.
 *
 * The cached result is exactly what Excel/LibreOffice last computed, so a static
 * PDF shows the expected number; cell styles (number formats, fills) are stored
 * separately and survive reassigning the value.
 */
export function flattenFormulas(ws: ExcelJS.Worksheet): void {
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.type === ExcelJS.ValueType.Formula) {
        // Read `cell.result` rather than `cell.value.result`: exceljs omits the
        // `result` key from the value object when it is falsy (0, '', false), so
        // pulling it off `value` would drop a real 0 and blank the cell.
        cell.value = cleanCellResult(cell.result);
      } else if (cell.type === ExcelJS.ValueType.Error) {
        cell.value = null;
      }
    });
  });
}

/**
 * Keep a real formula result; map only empties and errors to a blank cell.
 *
 * The checks are strict (`===`) on purpose: a result of `0` (or `false`) is a
 * genuine computed value and must be preserved. Only a missing result, a
 * spreadsheet error, or a formula that computes to an empty string `""` blanks.
 */
function cleanCellResult(result: unknown): ExcelJS.CellValue {
  // No cached result at all => render blank.
  if (result === undefined || result === null) return null;
  // exceljs models an error result as `{ error: '#NAME?' }`.
  if (typeof result === 'object' && result !== null && 'error' in result) {
    return null;
  }
  // A formula that computes to an empty string renders blank, but 0/false stay.
  if (result === '') return null;
  return result as ExcelJS.CellValue;
}

/**
 * Boolean font flags that exceljs reads with its `BooleanXform`, which treats an
 * element's mere *presence* as `true` and ignores its `val` attribute. A font
 * that explicitly turns one of these *off* (`<b val="false"/>`, `<i val="0"/>`,
 * `<strike val="false"/>`, …) — the form LibreOffice and several other producers
 * emit — is therefore misread as on, then re-serialized as a bare `<b/>`/`<i/>`/
 * `<strike/>`. LibreOffice honors that and the exported PDF shows bold, italic,
 * or strikethrough the cells never had. (Underline is parsed by a separate xform
 * that keeps `val`, so it is unaffected and not listed here.)
 */
const DISABLED_FONT_FLAGS = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow'];

/**
 * Matches a boolean font flag element whose value is explicitly false — i.e.
 * `<b val="false"/>`, `<i val='0'/>`, or the rare paired `<strike val="0"></strike>`
 * form — across either quote style and with optional surrounding whitespace. The
 * tag name is required to be one of {@link DISABLED_FONT_FLAGS}; genuine
 * `<b/>` / `<b val="true"/>` carry no false `val` and are left untouched.
 */
const DISABLED_FONT_FLAG_RE = new RegExp(
  `<(${DISABLED_FONT_FLAGS.join('|')})\\s+val=(["'])\\s*(?:0|false)\\s*\\2\\s*(?:/>|></\\1>)`,
  'gi',
);

/**
 * Strip explicitly-disabled boolean font flags out of an `xl/styles.xml`
 * document so exceljs cannot misread them as enabled (see
 * {@link DISABLED_FONT_FLAGS}). Operating on the raw XML — before exceljs parses
 * it — is what preserves real formatting: a genuine `<b/>` (bold actually on)
 * has no `val` and is kept, while `<b val="false"/>` (bold off) is removed, a
 * distinction that is already lost once exceljs collapses both to `bold: true`.
 */
export function sanitizeFontFlags(stylesXml: string): string {
  return stylesXml.replace(DISABLED_FONT_FLAG_RE, '');
}

/**
 * Read `sourcePath` and return its bytes with `xl/styles.xml` sanitized of
 * disabled font flags. xlsx is a zip, so the styles part is rewritten in place
 * and the archive repacked for `wb.xlsx.load`. Any failure to open or rewrite
 * the zip falls back to the original bytes, so a malformed or unexpected package
 * still exports rather than throwing.
 */
async function sanitizedWorkbookBuffer(sourcePath: string): Promise<Buffer> {
  const original = fs.readFileSync(path.resolve(sourcePath));
  try {
    const zip = await JSZip.loadAsync(original);
    const entry = zip.file('xl/styles.xml');
    if (!entry) return original;
    const xml = await entry.async('string');
    const sanitized = sanitizeFontFlags(xml);
    if (sanitized === xml) return original;
    zip.file('xl/styles.xml', sanitized);
    return await zip.generateAsync({ type: 'nodebuffer' });
  } catch {
    return original;
  }
}

/** True when the shell or OS could not locate the given command. */
function isCommandNotFound(err: unknown): boolean {
  const e = err as { code?: number | string; stderr?: string; message?: string };
  // Numeric exit codes: 127 = POSIX "command not found", 9009 = Windows cmd.exe
  if (e.code === 'ENOENT' || e.code === 127 || e.code === 9009) return true;
  // Windows puts the human-readable "not recognized" text in stderr/message
  // rather than relying on a stable exit code across all cmd.exe versions.
  const text = `${e.stderr ?? ''} ${e.message ?? ''}`;
  return /is not recognized as an internal or external command/i.test(text) ||
    /command not found/i.test(text);
}

/** Well-known default installation paths for LibreOffice per platform. */
const LIBREOFFICE_COMMON_PATHS: readonly string[] = (() => {
  switch (process.platform) {
    case 'win32':
      return [
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
      ];
    case 'darwin':
      return ['/Applications/LibreOffice.app/Contents/MacOS/soffice'];
    default:
      return [
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice',
      ];
  }
})();

/**
 * Return the absolute LibreOffice path written by the postinstall script
 * (`scripts/find-libreoffice.js`), or `null` if the cache is missing/stale.
 */
function cachedLibreOfficePath(): string | null {
  try {
    // Compiled layout: dist/output/actions.js → ../../.libreoffice-path
    const cachePath = path.join(__dirname, '..', '..', '.libreoffice-path');
    const p = fs.readFileSync(cachePath, 'utf8').trim();
    return p && fs.existsSync(p) ? p : null;
  } catch {
    return null;
  }
}

/**
 * Attempt each LibreOffice candidate in order, stopping on first success.
 *
 * Resolution order:
 *   1. Path cached by `npm install` (postinstall script)
 *   2. `libreoffice` / `soffice` via system PATH
 *   3. Common absolute installation paths (runtime fallback)
 *
 * Throws `LIBREOFFICE_NOT_FOUND` when none of the candidates work.
 */
async function spawnLibreOffice(sourcePath: string, outDir: string): Promise<void> {
  const cached   = cachedLibreOfficePath();
  const absolute = LIBREOFFICE_COMMON_PATHS.filter(
    (p) => fs.existsSync(p) && p !== cached,
  );
  const candidates = [
    ...(cached ? [cached] : []),
    'libreoffice',
    'soffice',
    ...absolute,
  ];

  for (const bin of candidates) {
    // Quote executables whose path contains spaces (common on Windows).
    const quoted = /\s/.test(bin) ? `"${bin}"` : bin;
    try {
      await execAsync(
        `${quoted} --headless --convert-to pdf --outdir "${outDir}" "${sourcePath}"`,
      );
      return;
    } catch (err) {
      if (isCommandNotFound(err)) continue;
      throw err;
    }
  }

  throw new ExtractError(
    'LIBREOFFICE_NOT_FOUND',
    'LibreOffice was not found on this system.',
    [
      'Install LibreOffice: https://www.libreoffice.org/download/libreoffice-still/',
      'Then run `npm install` to cache its path, or add it to your PATH.',
      `Default locations: ${LIBREOFFICE_COMMON_PATHS.join(', ')}`,
    ].join('\n  '),
  );
}

/**
 * Export `sourcePath` to PDF at `outputPath` using LibreOffice headless.
 *
 * LibreOffice always places its output as `<source-stem>.pdf` inside the
 * output directory. If that name differs from the requested `outputPath` the
 * file is renamed to match.
 */
async function convertWithLibreOffice(
  sourcePath: string,
  outputPath: string,
): Promise<void> {
  const absSource = path.resolve(sourcePath);
  const absOutput = path.resolve(outputPath);
  const outDir = path.dirname(absOutput);

  fs.mkdirSync(outDir, { recursive: true });
  await spawnLibreOffice(absSource, outDir);

  // Rename if LibreOffice's implicit <stem>.pdf differs from the requested path.
  const stem = path.basename(absSource, path.extname(absSource));
  const libreOut = path.join(outDir, `${stem}.pdf`);
  if (path.resolve(libreOut) !== absOutput) {
    fs.renameSync(libreOut, absOutput);
  }
}

/**
 * Produce an OS-appropriate variable assignment snippet. Multi-line values are
 * collapsed onto one line (joined with `; `) so they remain assignable.
 */
export function renderVarExport(
  name: string,
  text: string,
  config: ExtractConfig,
): string {
  const value = text.replace(/\r?\n/g, '; ').trim();
  const style =
    config.varExport === 'auto'
      ? process.platform === 'win32'
        ? 'cmd'
        : 'posix'
      : config.varExport;

  switch (style) {
    case 'cmd':
      return `set "${name}=${value}"`;
    case 'powershell':
      return `$env:${name} = "${value}"`;
    case 'posix':
    default:
      return `export ${name}="${value}"`;
  }
}
