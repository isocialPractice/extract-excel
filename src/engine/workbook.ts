/**
 * Workbook loading and the in-memory Sheet abstraction.
 *
 * Extraction logic never talks to exceljs directly — it operates on the simple
 * {@link Sheet} model below. That keeps cell/range/title extraction identical
 * whether the data came from an `.xlsx` workbook, a `.csv` file, or a raw CSV
 * string passed to `--test custom:"...."`.
 */
import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { CellAddress, RangeAddress, parseRange } from './address';
import { ExtractError, AmbiguousSheetError } from '../errors';
import { XmlMapping } from '../parser/types';

/** The XML-Schema-instance namespace Excel stamps on its XML-map exports. */
const XSI_NAMESPACE = 'http://www.w3.org/2001/XMLSchema-instance';

/** A single value-addressable grid. Cells are read by 1-based row/col. */
export class Sheet {
  readonly name: string;
  /** rows[r][c] holds the display value, 1-based indices mapped to 0-based. */
  private readonly rows: Array<Array<unknown>>;
  /** Merged-cell regions (e.g. a header spanning A1:N1), 1-based. */
  readonly merges: RangeAddress[];

  constructor(name: string, rows: Array<Array<unknown>>, merges: RangeAddress[] = []) {
    this.name = name;
    this.rows = rows;
    this.merges = merges;
  }

  /** Number of populated rows. */
  get rowCount(): number {
    return this.rows.length;
  }

  /**
   * Compute the used range: the bounding box of every non-empty cell. Returns
   * null when the sheet is entirely empty. Powers the `--range sheet` preset.
   */
  usedRange(): RangeAddress | null {
    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -Infinity;
    let maxCol = -Infinity;
    for (let r = 0; r < this.rows.length; r++) {
      const row = this.rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const v = normalizeValue(row[c]);
        if (v !== null && v !== '') {
          if (r + 1 < minRow) minRow = r + 1;
          if (c + 1 < minCol) minCol = c + 1;
          if (r + 1 > maxRow) maxRow = r + 1;
          if (c + 1 > maxCol) maxCol = c + 1;
        }
      }
    }
    if (maxRow < 0) return null;
    return { start: { row: minRow, col: minCol }, end: { row: maxRow, col: maxCol } };
  }

  /**
   * Read a cell's display value. Out-of-bounds reads return `null`, matching a
   * spreadsheet's "empty cell" semantics rather than throwing.
   */
  getValue(addr: CellAddress): string | null {
    const row = this.rows[addr.row - 1];
    if (!row) return null;
    const value = row[addr.col - 1];
    return normalizeValue(value);
  }

  /** True when a cell holds no meaningful data. */
  isEmpty(addr: CellAddress): boolean {
    const v = this.getValue(addr);
    return v === null || v === '';
  }
}

/** A named collection of sheets. */
export class Workbook {
  readonly source: string;
  private readonly sheets: Sheet[];

  constructor(source: string, sheets: Sheet[]) {
    this.source = source;
    this.sheets = sheets;
  }

  get sheetNames(): string[] {
    return this.sheets.map((s) => s.name);
  }

  get sheetCount(): number {
    return this.sheets.length;
  }

  /**
   * Resolve the sheet to extract from.
   *
   * Per the doc's IMPORTANT note: if no sheet is named and the workbook has
   * more than one sheet, this is a custom error. With exactly one sheet, the
   * selection is unambiguous and `name` is ignored when absent.
   */
  resolveSheet(name?: string): Sheet {
    if (name !== undefined) {
      const found = this.sheets.find(
        (s) => s.name.toLowerCase() === name.toLowerCase(),
      );
      if (!found) {
        throw new ExtractError(
          'SHEET_NOT_FOUND',
          `Sheet "${name}" not found in "${this.source}".`,
          `Available sheets: ${this.sheetNames.map((s) => `"${s}"`).join(', ')}`,
        );
      }
      return found;
    }
    if (this.sheets.length === 0) {
      throw new ExtractError('UNREADABLE_WORKBOOK', `"${this.source}" has no sheets.`);
    }
    if (this.sheets.length > 1) {
      throw new AmbiguousSheetError(this.source, this.sheetNames);
    }
    return this.sheets[0];
  }
}

/** Coerce an exceljs / CSV cell value into a trimmed string (or null). */
function normalizeValue(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value instanceof Date) return value.toISOString();
  // exceljs rich text / formula / hyperlink objects.
  const obj = value as Record<string, unknown>;
  if (typeof obj.text === 'string') return obj.text;
  if (typeof obj.result !== 'undefined') return normalizeValue(obj.result);
  if (Array.isArray(obj.richText)) {
    return obj.richText.map((rt) => (rt as { text?: string }).text ?? '').join('');
  }
  if (typeof obj.hyperlink === 'string') return obj.hyperlink;
  // Spreadsheet error objects (`{ error: '#REF!' }`) and any other cell-value
  // shape we don't recognise carry no usable text — treat them as empty rather
  // than stringifying to the useless "[object Object]".
  if (typeof obj.error === 'string') return obj.error;
  return null;
}

/** Build a {@link Sheet} from a 2-D array (used for CSV and raw data). */
function sheetFromGrid(name: string, grid: Array<Array<unknown>>): Sheet {
  return new Sheet(name, grid);
}

/** Convert an exceljs worksheet to our {@link Sheet}, preserving merges. */
function sheetFromExcelJs(ws: ExcelJS.Worksheet): Sheet {
  const grid: Array<Array<unknown>> = [];
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const cells: unknown[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cells[colNumber - 1] = cell.value;
    });
    grid[rowNumber - 1] = cells;
  });
  return new Sheet(ws.name, grid, readMerges(ws));
}

/** Read merged-cell ranges from an exceljs worksheet (e.g. `["A1:N1"]`). */
function readMerges(ws: ExcelJS.Worksheet): RangeAddress[] {
  const model = (ws as unknown as { model?: { merges?: string[] } }).model;
  const merges = model?.merges ?? [];
  const out: RangeAddress[] = [];
  for (const ref of merges) {
    try {
      out.push(parseRange(ref));
    } catch {
      // Ignore any merge ref we cannot parse rather than failing the load.
    }
  }
  return out;
}

/**
 * Parse RFC-4180-ish CSV text into a grid. Handles quoted fields, escaped
 * quotes (`""`), and embedded commas/newlines. Sufficient for "read this as if
 * opened as a CSV in Excel" per the doc.
 */
export function parseCsv(text: string): Array<Array<string>> {
  const rows: string[][] = [];
  let field = '';
  let row: string[] = [];
  let inQuotes = false;
  // Normalize literal "\n" escape sequences that arrive from shell test args.
  const src = text.replace(/\\r\\n|\\n/g, '\n').replace(/\r\n/g, '\n');

  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (inQuotes) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(field.trim());
      field = '';
    } else if (ch === '\n') {
      row.push(field.trim());
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += ch;
    }
  }
  // Flush the trailing field/row if the text did not end with a newline.
  if (field.length > 0 || row.length > 0) {
    row.push(field.trim());
    rows.push(row);
  }
  return rows;
}

/** Build a Workbook from raw CSV text (single sheet). */
export function workbookFromCsvText(text: string, name = 'Sheet1'): Workbook {
  const grid = parseCsv(text);
  return new Workbook(`<raw:${name}>`, [sheetFromGrid(name, grid)]);
}

/** Load a workbook from a `.xlsx`/`.xlsm` file or a `.csv` file on disk. */
export async function loadWorkbook(filePath: string): Promise<Workbook> {
  const resolved = path.resolve(filePath);
  if (!fs.existsSync(resolved)) {
    throw new ExtractError(
      'WORKBOOK_NOT_FOUND',
      `Workbook "${filePath}" was not found.`,
      `Looked relative to the current working directory: ${resolved}`,
    );
  }

  const ext = path.extname(resolved).toLowerCase();
  try {
    if (ext === '.csv') {
      const text = fs.readFileSync(resolved, 'utf8');
      const grid = parseCsv(text);
      return new Workbook(filePath, [sheetFromGrid(path.basename(resolved, ext), grid)]);
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(resolved);
    const sheets: Sheet[] = [];
    wb.eachSheet((ws) => sheets.push(sheetFromExcelJs(ws)));
    return new Workbook(filePath, sheets);
  } catch (err) {
    if (err instanceof ExtractError) throw err;
    throw new ExtractError(
      'UNREADABLE_WORKBOOK',
      `Could not read "${filePath}": ${(err as Error).message}`,
    );
  }
}

/**
 * Read the Excel XML-map container tags from a workbook, if it has one.
 *
 * Workbooks built with Excel's "XML Source" mapping feature embed the schema at
 * `xl/xmlMaps.xml`. exceljs does not expose it, so the file is reopened as a zip
 * (xlsx is a zip) and the map is parsed directly. Returns the root + repeating
 * (row) element names plus the `xmlns:xsi` namespace that mirrors Excel's own
 * XML export. Returns `undefined` for any workbook without a usable map — the
 * common case — so callers fall back to the sheet-name defaults.
 *
 * Never throws: a missing entry, a non-zip file, or an unparsable map all yield
 * `undefined` rather than failing the extraction.
 */
export async function readXmlMapping(filePath: string): Promise<XmlMapping | undefined> {
  try {
    const buffer = fs.readFileSync(path.resolve(filePath));
    const zip = await JSZip.loadAsync(buffer);
    const entry = zip.file('xl/xmlMaps.xml');
    if (!entry) return undefined;
    const parsed = parseXmlMap(await entry.async('string'));
    if (!parsed) return undefined;
    return { ...parsed, namespaces: { 'xmlns:xsi': XSI_NAMESPACE } };
  } catch {
    return undefined;
  }
}

/**
 * Parse the root and repeating-row element names out of an `xl/xmlMaps.xml`
 * document.
 *
 * The root comes from the `<Map RootElement="...">` attribute; the repeating
 * row element is the schema element declared `maxOccurs="unbounded"`. Both
 * lookups tolerate the `xsd:`/`xs:` prefix variants and either attribute order.
 * Returns `null` when the document does not yield both names.
 */
export function parseXmlMap(xml: string): { root: string; row: string } | null {
  // Every `<element ...>` declaration in the embedded schema, with its attrs.
  const elements: Array<{ name?: string; unbounded: boolean }> = [];
  const elementRe = /<(?:[A-Za-z][\w.-]*:)?element\b([^>]*?)\/?>/gi;
  let match: RegExpExecArray | null;
  while ((match = elementRe.exec(xml)) !== null) {
    const attrs = match[1];
    elements.push({
      name: /\bname\s*=\s*"([^"]+)"/i.exec(attrs)?.[1],
      unbounded: /\bmaxOccurs\s*=\s*"unbounded"/i.test(attrs),
    });
  }

  // Root: prefer the Map's RootElement attribute, else the first declared element
  // (the schema's document element).
  const root =
    /<Map\b[^>]*\bRootElement\s*=\s*"([^"]+)"/i.exec(xml)?.[1] ??
    elements.find((e) => e.name)?.name;
  // Row: the element marked as repeating (maxOccurs="unbounded").
  const row = elements.find((e) => e.unbounded && e.name)?.name;

  return root && row ? { root, row } : null;
}
