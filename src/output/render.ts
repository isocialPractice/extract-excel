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

/** Infer a sensible format from a file path's extension. */
export function formatFromPath(filePath: string): OutputFormat | null {
  const lower = filePath.toLowerCase();
  if (lower.endsWith('.csv')) return 'csv';
  if (lower.endsWith('.md') || lower.endsWith('.markdown')) return 'markdown';
  return null;
}
