/**
 * Renderers: turn structured {@link ExtractionResult}s into output text in a
 * chosen format.
 *
 *   - text     : the documented terminal layout (column separator, blank line
 *                between blocks, trailing-empty columns trimmed per row).
 *   - csv      : RFC-4180 CSV write-back (quoting as needed).
 *   - table    : aligned, bordered monospace grid.
 *   - markdown : GitHub-style table; merged cells are spanned by repeating the
 *                master value so the grid stays as close to Excel as possible.
 *   - xml      : a field-mapped XML document — the sheet name becomes the root
 *                element (camel-cased), the first extracted row supplies the
 *                field names, and every later row becomes a `<row>` record of
 *                those named fields.
 *   - aligned  : the PDF/`.md` layout — a monospace, pipe-bordered table whose
 *                columns are padded to line up (two-pass: measure, then emit).
 */
import { ExtractionResult } from '../engine/extract';
import { ExtractConfig } from '../config';
import { OutputFormat } from '../parser/types';
import { RangeAddress } from '../engine/address';

/** Render results to the requested format. */
export function render(
  results: ExtractionResult[],
  format: OutputFormat,
  config: ExtractConfig,
): string {
  switch (format) {
    case 'csv':
      return renderCsv(results);
    case 'table':
      return renderTable(results);
    case 'markdown':
      return renderMarkdown(results);
    case 'xml':
      return renderXml(results);
    case 'text':
    default:
      return renderText(results, config);
  }
}

/** The documented terminal layout. */
export function renderText(results: ExtractionResult[], config: ExtractConfig): string {
  const blockSep = `\n${config.blockSeparator}\n`;
  return results
    .map((result) =>
      result.rows
        .map((row) => trimTrailingEmpty(row).join(config.columnSeparator))
        .join('\n'),
    )
    .join(blockSep);
}

/** CSV write-back; multiple extract blocks are separated by a blank line. */
export function renderCsv(results: ExtractionResult[]): string {
  return results
    .map((result) => result.rows.map((row) => row.map(csvField).join(',')).join('\n'))
    .join('\n\n');
}

/** Aligned, bordered monospace table. */
export function renderTable(results: ExtractionResult[]): string {
  return results
    .map((result) => {
      const grid = spanMerges(result);
      if (grid.length === 0) return '';
      const widths = columnWidths(grid);
      const bar = '+' + widths.map((w) => '-'.repeat(w + 2)).join('+') + '+';
      const lines = [bar];
      grid.forEach((row, i) => {
        const cells = widths.map(
          (w, c) => ' ' + (row[c] ?? '').padEnd(w) + ' ',
        );
        lines.push('|' + cells.join('|') + '|');
        if (i === 0) lines.push(bar); // header rule after the first row
      });
      lines.push(bar);
      return lines.join('\n');
    })
    .join('\n\n');
}

/** GitHub-flavored markdown table with merged-cell spanning. */
export function renderMarkdown(results: ExtractionResult[]): string {
  return results
    .map((result) => {
      const grid = spanMerges(result);
      if (grid.length === 0) return '';
      const colCount = Math.max(...grid.map((r) => r.length));
      const pad = (row: string[]): string[] =>
        Array.from({ length: colCount }, (_, c) => mdField(row[c] ?? ''));
      const header = pad(grid[0]);
      const lines = [
        `| ${header.join(' | ')} |`,
        `| ${header.map(() => '---').join(' | ')} |`,
      ];
      for (let i = 1; i < grid.length; i++) {
        lines.push(`| ${pad(grid[i]).join(' | ')} |`);
      }
      return lines.join('\n');
    })
    .join('\n\n');
}

/**
 * Field-mapped XML document of the extracted results.
 *
 * Each {@link ExtractionResult} is treated as a table whose first row is the
 * header: the sheet name (camel-cased) becomes the root element, the header
 * cells become the per-field element names (non-alphanumeric runs collapse to
 * `_`, so `Last Name` => `<Last_Name>` and `FT/PT` => `<FT_PT>`), and every
 * subsequent row is emitted as a `<row>` record of those named fields.
 *
 * A single result (the documented `--xml` case — one sheet) is rendered with
 * the sheet element as the document root. When several results are present they
 * are nested under a synthetic `<extract>` root so the document stays
 * well-formed with its single required root element.
 */
export function renderXml(results: ExtractionResult[]): string {
  const lines = ['<?xml version="1.0" encoding="UTF-8"?>'];
  if (results.length <= 1) {
    lines.push(...xmlSheet(results[0], 0));
  } else {
    lines.push('<extract>');
    for (const result of results) lines.push(...xmlSheet(result, 1));
    lines.push('</extract>');
  }
  return lines.join('\n');
}

/**
 * Render one result as a record list.
 *
 * Container tags come from the result's resolved {@link ExtractionResult.xmlMapping}
 * (Excel XML map / `--xml:root,row` override): `root` is the document element
 * and `row` the repeating record element, with any namespace attributes on the
 * root. Without a mapping it falls back to the camel-cased sheet name and a
 * `<row>` element.
 */
function xmlSheet(result: ExtractionResult | undefined, depth: number): string[] {
  const pad = '  '.repeat(depth);
  const mapping = result?.xmlMapping;
  const root = mapping?.root ?? camelCaseName(result?.sheetName ?? 'data');
  const rowTag = mapping?.row ?? 'row';
  const attrs = mapping?.namespaces
    ? Object.entries(mapping.namespaces)
        .map(([key, value]) => ` ${key}="${xmlAttr(value)}"`)
        .join('')
    : '';
  const rows = result?.rows ?? [];
  if (rows.length === 0) return [`${pad}<${root}${attrs}/>`];

  // The first extracted row is the header; it names the fields, it is not data.
  const fields = rows[0].map((header, i) => sanitizeFieldName(header, i));
  const out = [`${pad}<${root}${attrs}>`];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    // Skip rows with no data in any mapped field (blank separators, and the
    // junk rows that used to surface as "[object Object]" before normalization).
    if (fields.every((_, c) => (row[c] ?? '') === '')) continue;
    out.push(`${pad}  <${rowTag}>`);
    fields.forEach((tag, c) => {
      out.push(`${pad}    <${tag}>${xmlText(row[c] ?? '')}</${tag}>`);
    });
    out.push(`${pad}  </${rowTag}>`);
  }
  out.push(`${pad}</${root}>`);
  return out;
}

/**
 * Camel-case a sheet name into a valid XML root element name: `Data` => `data`,
 * `On Boarding` => `onBoarding`. A leading digit is prefixed with `_` so the
 * result is a legal XML name; an empty name falls back to `data`.
 */
function camelCaseName(name: string): string {
  const words = name.split(/[^A-Za-z0-9]+/).filter(Boolean);
  if (words.length === 0) return 'data';
  const head = words[0].charAt(0).toLowerCase() + words[0].slice(1);
  const tail = words.slice(1).map((w) => w.charAt(0).toUpperCase() + w.slice(1));
  const out = [head, ...tail].join('');
  return /^[0-9]/.test(out) ? `_${out}` : out;
}

/**
 * Turn a header cell into a valid XML element name: non-alphanumeric runs become
 * a single `_`, a leading digit is prefixed with `_`, and a blank header falls
 * back to a positional `column_N` (1-based) so every field is addressable.
 */
function sanitizeFieldName(header: string, index: number): string {
  const cleaned = header.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  if (cleaned === '') return `column_${index + 1}`;
  return /^[0-9]/.test(cleaned) ? `_${cleaned}` : cleaned;
}

/**
 * Aligned, pipe-bordered monospace table — the format used by the `pdf` and
 * `md` output targets.
 *
 * The layout is built in two passes, as the spec describes:
 *   1. Measure: find the widest cell in each column so columns can line up.
 *   2. Emit: pad every cell to its column width and write the `| … | … |` rows.
 *
 * Two structural touches make real-world sheets read cleanly:
 *   - A fully empty row splits the result into independent sub-tables (each with
 *     its own columns and widths), so a summary block below a blank row renders
 *     as its own aligned table rather than being stretched to the main grid.
 *   - Columns that are empty across a whole sub-table are dropped.
 *
 * When the result's `assumeMerge` flag is set, adjacent duplicate *text* cells
 * (the artifacts of a spanned merge, e.g. `Total Employees | Total Employees`)
 * are collapsed back into a single cell; purely numeric duplicates are kept.
 *
 * Only the first sub-table receives a markdown header rule (`|---|---|`) under
 * its first row; later sub-tables are continuation data with no header.
 */
export function renderAligned(results: ExtractionResult[]): string {
  return results
    .map((result) => renderAlignedResult(result))
    .filter((text) => text.length > 0)
    .join('\n\n');
}

/** Render one extraction result as one or more aligned sub-tables. */
function renderAlignedResult(result: ExtractionResult): string {
  const grid = spanMerges(result);
  const blocks = splitOnEmptyRows(grid);
  const tables: string[] = [];

  blocks.forEach((block, blockIndex) => {
    let rows = result.assumeMerge ? block.map(collapseDuplicateText) : block;
    rows = dropEmptyColumns(padRectangular(rows));
    if (rows.length === 0) return;

    const widths = columnWidths(rows);
    const lines: string[] = [];
    rows.forEach((row, rowIndex) => {
      lines.push(alignedRow(row, widths));
      // Only the very first sub-table's first row is treated as a header.
      if (blockIndex === 0 && rowIndex === 0) lines.push(alignedRule(widths));
    });
    tables.push(lines.join('\n'));
  });

  return tables.join('\n\n');
}

/** Format one row: `| a | b | c |`, each cell padded to its column width. */
function alignedRow(row: string[], widths: number[]): string {
  const cells = widths.map((w, c) => (row[c] ?? '').padEnd(w));
  return `| ${cells.join(' | ')} |`;
}

/** Format the header rule: `|-----|-----|`, dashes filling each column box. */
function alignedRule(widths: number[]): string {
  return '|' + widths.map((w) => '-'.repeat(w + 2)).join('|') + '|';
}

/** Split a grid into sub-grids separated by one or more fully empty rows. */
function splitOnEmptyRows(grid: string[][]): string[][][] {
  const blocks: string[][][] = [];
  let current: string[][] = [];
  for (const row of grid) {
    if (row.every((cell) => cell.trim() === '')) {
      if (current.length > 0) {
        blocks.push(current);
        current = [];
      }
    } else {
      current.push(row);
    }
  }
  if (current.length > 0) blocks.push(current);
  return blocks;
}

/** Pad every row to the grid's widest row length with empty strings. */
function padRectangular(rows: string[][]): string[][] {
  const width = rows.reduce((max, r) => Math.max(max, r.length), 0);
  return rows.map((r) =>
    r.length === width ? r : [...r, ...Array(width - r.length).fill('')],
  );
}

/** Remove columns that are empty in every row of the (rectangular) grid. */
function dropEmptyColumns(rows: string[][]): string[][] {
  if (rows.length === 0) return rows;
  const width = rows[0].length;
  const keep: number[] = [];
  for (let c = 0; c < width; c++) {
    if (rows.some((row) => (row[c] ?? '').trim() !== '')) keep.push(c);
  }
  return rows.map((row) => keep.map((c) => row[c] ?? ''));
}

/**
 * A cell is "numeric" when it is built solely from digits and the numeric
 * punctuation called out in the spec (`/ + * - . \` and `,`). Such duplicates
 * are real data and left alone; non-numeric (text) duplicates are merge
 * artifacts and collapsed by {@link collapseDuplicateText}.
 */
function isNumericish(value: string): boolean {
  return value !== '' && /^[0-9/+*.\\,-]+$/.test(value);
}

/** Collapse runs of adjacent, identical, non-numeric cells down to one. */
function collapseDuplicateText(row: string[]): string[] {
  const out: string[] = [];
  for (const cell of row) {
    const prev = out[out.length - 1];
    if (out.length > 0 && cell === prev && cell !== '' && !isNumericish(cell)) {
      continue; // duplicate text cell => assumed merge artifact, drop it
    }
    out.push(cell);
  }
  return out;
}

/**
 * Apply merged-cell spans: fill every non-master cell of a merge with the
 * master (top-left) value. Markdown/aligned grids cannot truly span, so a
 * spanned header reads as the same value repeated across its columns/rows —
 * the closest faithful representation of the Excel layout.
 */
function spanMerges(result: ExtractionResult): string[][] {
  const rows = result.rows.map((r) => [...r]);
  const merges = result.merges ?? [];
  for (const m of merges) {
    const master = rows[m.start.row - 1]?.[m.start.col - 1] ?? '';
    for (let r = m.start.row; r <= m.end.row; r++) {
      for (let c = m.start.col; c <= m.end.col; c++) {
        if (r === m.start.row && c === m.start.col) continue;
        const row = rows[r - 1];
        if (row) row[c - 1] = master;
      }
    }
  }
  return rows;
}

/** Drop trailing empty strings from a row of columns. */
function trimTrailingEmpty(cols: string[]): string[] {
  let end = cols.length;
  while (end > 0 && cols[end - 1] === '') end--;
  return cols.slice(0, end);
}

/** Compute the display width of each column across a grid. */
function columnWidths(grid: string[][]): number[] {
  const colCount = Math.max(...grid.map((r) => r.length));
  const widths = new Array<number>(colCount).fill(0);
  for (const row of grid) {
    for (let c = 0; c < colCount; c++) {
      const len = (row[c] ?? '').length;
      if (len > widths[c]) widths[c] = len;
    }
  }
  return widths;
}

/** Quote a CSV field when it contains a comma, quote, or newline. */
function csvField(value: string): string {
  if (/[",\n\r]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

/** Escape pipe characters and collapse newlines for a markdown cell. */
function mdField(value: string): string {
  return value.replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
}

/** Escape the XML-significant characters in element text content. */
function xmlText(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Escape an XML attribute value (text escaping plus the double quote). */
function xmlAttr(value: string): string {
  return xmlText(value).replace(/"/g, '&quot;');
}

/** Infer a sensible format from a file path's extension. */
export function formatFromPath(filePath: string): OutputFormat | null {
  const lower = filePath.toLowerCase();
  if (lower.endsWith('.csv')) return 'csv';
  if (lower.endsWith('.md') || lower.endsWith('.markdown')) return 'markdown';
  if (lower.endsWith('.xml')) return 'xml';
  return null;
}
