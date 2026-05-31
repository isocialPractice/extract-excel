/**
 * Core extraction primitives: cell, range, and title.
 *
 * Each function returns a structured {@link ExtractionResult} (rows of column
 * values). Rendering to text lives in `output/format.ts`, so the same result
 * can be sent to stdout, a file, a PDF, or a shell variable unchanged.
 */
import { Sheet } from './workbook';
import {
  parseCell,
  parseRange,
  formatCell,
  formatRange,
  CellAddress,
  RangeAddress,
} from './address';
import { ExtractError } from '../errors';
import { ExtractConfig } from '../config';

export type ExtractionKind = 'cell' | 'range' | 'title';

export interface ExtractionResult {
  kind: ExtractionKind;
  /** Canonical reference/label of what was extracted (e.g. `YC30`, `Address`). */
  ref: string;
  /**
   * Output grid: one inner array per line, holding that line's columns. For
   * ranges this is a full rectangle (no trimming) so structured renderers
   * (csv/table/markdown) see every column; the text renderer trims trailing
   * empties to match the documented terminal output.
   */
  rows: string[][];
  /**
   * Merged-cell regions intersecting a range, expressed relative to the result
   * grid (1-based). Used by table/markdown/PDF renderers to span values.
   */
  merges?: RangeAddress[];
}

/** Extract a single cell's value. */
export function extractCell(sheet: Sheet, ref: string): ExtractionResult {
  const addr = parseCell(ref);
  const value = sheet.getValue(addr) ?? '';
  return { kind: 'cell', ref: formatCell(addr), rows: [[value]] };
}

/** Extract a rectangular range from an A1 reference (e.g. `A2:F45`). */
export function extractRange(sheet: Sheet, ref: string): ExtractionResult {
  return extractRangeByAddress(sheet, parseRange(ref));
}

/**
 * Extract a rectangular range by address. Produces a full rectangle plus any
 * merged regions (clipped and re-based to the grid), leaving trailing-empty
 * trimming to the text renderer.
 */
export function extractRangeByAddress(sheet: Sheet, range: RangeAddress): ExtractionResult {
  const rows: string[][] = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    const cols: string[] = [];
    for (let c = range.start.col; c <= range.end.col; c++) {
      cols.push(sheet.getValue({ row: r, col: c }) ?? '');
    }
    rows.push(cols);
  }
  return {
    kind: 'range',
    ref: formatRange(range),
    rows,
    merges: clipMerges(sheet.merges, range),
  };
}

/**
 * Extract a whole sheet's used range (the `--range sheet` preset): the bounding
 * box from the first to the last cell containing data.
 */
export function extractUsedRange(sheet: Sheet): ExtractionResult {
  const range = sheet.usedRange();
  if (!range) {
    return { kind: 'range', ref: '(empty)', rows: [], merges: [] };
  }
  return extractRangeByAddress(sheet, range);
}

/** Clip sheet merges to a range and re-base them to grid coordinates. */
function clipMerges(merges: RangeAddress[], range: RangeAddress): RangeAddress[] {
  const out: RangeAddress[] = [];
  for (const m of merges) {
    const r1 = Math.max(m.start.row, range.start.row);
    const c1 = Math.max(m.start.col, range.start.col);
    const r2 = Math.min(m.end.row, range.end.row);
    const c2 = Math.min(m.end.col, range.end.col);
    if (r1 <= r2 && c1 <= c2) {
      out.push({
        start: { row: r1 - range.start.row + 1, col: c1 - range.start.col + 1 },
        end: { row: r2 - range.start.row + 1, col: c2 - range.start.col + 1 },
      });
    }
  }
  return out;
}

/**
 * Extract a column by its header title.
 *
 * The header row is the first non-empty row by default, or an explicit row when
 * `row=` was supplied (e.g. `--title row=2 "March"`). Values are collected from
 * the row below the header down to the first empty cell, matching the doc:
 * "extracts cells ... up-until the first empty cell".
 */
export function extractTitle(
  sheet: Sheet,
  title: string,
  options: { headerRow?: number; config: ExtractConfig },
): ExtractionResult {
  const headerRow = options.headerRow ?? findHeaderRow(sheet);
  const col = findTitleColumn(sheet, headerRow, title);
  if (col === null) {
    throw new ExtractError(
      'TITLE_NOT_FOUND',
      `Column title "${title}" was not found on row ${headerRow} of sheet "${sheet.name}".`,
      'Use --title row=<n> "<title>" to point at the correct header row.',
    );
  }

  const rows: string[][] = [];
  for (let r = headerRow + 1; r <= sheet.rowCount; r++) {
    const addr: CellAddress = { row: r, col };
    if (sheet.isEmpty(addr)) break; // stop at the first gap in the column
    rows.push([sheet.getValue(addr) ?? '']);
  }
  return { kind: 'title', ref: title, rows };
}

/** Locate the first row that contains any data. */
function findHeaderRow(sheet: Sheet): number {
  for (let r = 1; r <= sheet.rowCount; r++) {
    for (let c = 1; ; c++) {
      const v = sheet.getValue({ row: r, col: c });
      if (v === null) break; // reached the end of populated columns in this row
      if (v !== '') return r;
    }
  }
  return 1;
}

/** Find the column index on `headerRow` whose value matches `title`. */
function findTitleColumn(sheet: Sheet, headerRow: number, title: string): number | null {
  const target = title.trim().toLowerCase();
  for (let c = 1; ; c++) {
    const v = sheet.getValue({ row: headerRow, col: c });
    if (v === null && c > 1) break; // past the populated header columns
    if (v !== null && v.trim().toLowerCase() === target) return c;
    // Guard against unbounded scanning on a totally empty header row.
    if (c > 16384) break;
  }
  return null;
}
