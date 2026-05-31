/**
 * A1-notation address parsing and normalization.
 *
 * The draft doc deliberately uses messy, lower/upper-mixed references such as
 * `yc30`, `r249`, `y10:AA123`, and `B1:v28` to show the tool must normalize
 * them. All addresses here are 1-based (row 1, column 1 == A1), matching how
 * spreadsheet users think.
 */
import { ExtractError } from '../errors';

export interface CellAddress {
  /** 1-based row index. */
  row: number;
  /** 1-based column index (A => 1). */
  col: number;
}

export interface RangeAddress {
  start: CellAddress;
  end: CellAddress;
}

const CELL_RE = /^([A-Za-z]+)(\d+)$/;

/** Convert a column label (`A`, `Z`, `AA`, `yc`) to a 1-based index. */
export function columnToNumber(label: string): number {
  let n = 0;
  const upper = label.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i) - 64; // 'A' => 1
    if (code < 1 || code > 26) {
      throw new ExtractError('INVALID_CELL', `Invalid column label "${label}".`);
    }
    n = n * 26 + code;
  }
  return n;
}

/** Convert a 1-based column index back to its label (1 => `A`, 27 => `AA`). */
export function numberToColumn(num: number): string {
  if (num < 1 || !Number.isInteger(num)) {
    throw new ExtractError('INVALID_CELL', `Invalid column number "${num}".`);
  }
  let label = '';
  let n = num;
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

/** Parse a single cell reference into a normalized {row, col}. */
export function parseCell(ref: string): CellAddress {
  const match = CELL_RE.exec(ref.trim());
  if (!match) {
    throw new ExtractError(
      'INVALID_CELL',
      `"${ref}" is not a valid cell reference (expected e.g. A1, BA2, yc30).`,
    );
  }
  const col = columnToNumber(match[1]);
  const row = parseInt(match[2], 10);
  if (row < 1) {
    throw new ExtractError('INVALID_CELL', `Row in "${ref}" must be >= 1.`);
  }
  return { row, col };
}

/** Render a {row, col} back to canonical A1 form (`YC30`). */
export function formatCell(addr: CellAddress): string {
  return `${numberToColumn(addr.col)}${addr.row}`;
}

/**
 * Parse a range reference (`A2:F45`). The two corners are normalized so that
 * `start` is always the top-left and `end` the bottom-right, regardless of the
 * order they were written in.
 */
export function parseRange(ref: string): RangeAddress {
  const parts = ref.trim().split(':');
  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    throw new ExtractError(
      'INVALID_RANGE',
      `"${ref}" is not a valid range (expected e.g. A2:F45).`,
    );
  }
  const a = parseCell(parts[0]);
  const b = parseCell(parts[1]);
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

/** Render a range back to canonical A1 form (`A2:F45`). */
export function formatRange(range: RangeAddress): string {
  return `${formatCell(range.start)}:${formatCell(range.end)}`;
}
