/**
 * Public API for `extract-excel`.
 *
 * The same engine that powers the CLI is exposed here so external scripts,
 * frameworks, and npm projects can extract Excel data programmatically:
 *
 * ```ts
 * import { extract, loadWorkbook, extractRange } from 'extract-excel';
 *
 * // High-level: run a CLI-style command and capture the text.
 * const text = await extract(['book.xlsx', '--range', 'A1:C5']);
 *
 * // Low-level: drive the engine directly.
 * const wb = await loadWorkbook('book.xlsx');
 * const sheet = wb.resolveSheet('Sheet1');
 * const result = extractRange(sheet, 'A1:C5');
 * ```
 */
import { parse, splitArgString } from './parser/tokenizer';
import { runExtract, RunOptions } from './run';
import { runTests, TestReport } from './test/runner';
import { renderHelp } from './help/docs';
import { getVersion } from './version';
import { formatResults } from './output/format';
import { loadConfig } from './config';
import { Workbook, Sheet, loadWorkbook, workbookFromCsvText } from './engine/workbook';
import { extractCell, extractRange, extractTitle, ExtractionResult } from './engine/extract';
import { Command } from './parser/types';

export { parse, splitArgString } from './parser/tokenizer';
export { runExtract } from './run';
export type { RunOptions } from './run';
export { runTests } from './test/runner';
export type { TestReport, TestCaseResult } from './test/runner';
export { renderHelp } from './help/docs';
export { getVersion } from './version';
export { formatResult, formatResults } from './output/format';
export {
  render,
  renderText,
  renderCsv,
  renderTable,
  renderMarkdown,
  renderXml,
  renderAligned,
  formatFromPath,
} from './output/render';
export { dispatch, renderVarExport } from './output/actions';
export { loadConfig, DEFAULT_CONFIG } from './config';
export type { ExtractConfig } from './config';
export {
  Workbook,
  Sheet,
  loadWorkbook,
  workbookFromCsvText,
  parseCsv,
  readXmlMapping,
  parseXmlMap,
} from './engine/workbook';
export {
  extractCell,
  extractRange,
  extractRangeByAddress,
  extractUsedRange,
  extractTitle,
} from './engine/extract';
export type { ExtractionResult, ExtractionKind } from './engine/extract';
export * from './engine/address';
export * from './parser/types';
export { ExtractError, AmbiguousSheetError } from './errors';
export type { ExtractErrorCode } from './errors';

/**
 * Run a CLI-style argument array and return the text that would have been
 * written to stdout. Output targets such as files/PDFs still take effect; only
 * the terminal stream is captured and returned.
 *
 * Help and test commands are also supported: help returns its text, and tests
 * return a formatted summary (use {@link runTests} for the structured report).
 */
export async function extract(
  argv: string[],
  options: RunOptions = {},
): Promise<string> {
  const command: Command = parse(argv);

  if (command.kind === 'help') {
    return renderHelp(command);
  }
  if (command.kind === 'version') {
    return getVersion();
  }
  if (command.kind === 'test') {
    const report: TestReport = await runTests(command, { cwd: options.cwd });
    return report.output ?? formatTestSummary(report);
  }

  let captured = '';
  const passthrough = options.write;
  await runExtract(command, {
    ...options,
    write: (text) => {
      captured += text;
      if (passthrough) passthrough(text);
    },
  });
  return captured.replace(/\n$/, '');
}

/** Convenience: extract a single cell from a workbook file. */
export async function extractCellFromFile(
  file: string,
  ref: string,
  sheet?: string,
): Promise<string> {
  const wb = await loadWorkbook(file);
  const result = extractCell(wb.resolveSheet(sheet), ref);
  return formatResults([result], loadConfig());
}

/** Build a one-line, human-readable summary of a test report. */
export function formatTestSummary(report: TestReport): string {
  const lines = report.results.map(
    (r) => `  ${r.ok ? 'PASS' : 'FAIL'}  ${r.name}${r.error ? ` — ${r.error}` : ''}`,
  );
  lines.push('');
  lines.push(`  ${report.passed} passed, ${report.failed} failed`);
  return lines.join('\n');
}

/** Ergonomic alias for an extraction result. */
export type Extraction = ExtractionResult;
