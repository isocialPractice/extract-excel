/**
 * Execution layer: run a parsed {@link ExtractCommand} end-to-end.
 *
 * This module is the bridge between the tokenizer's data model and the engine +
 * output layers. It is intentionally side-effect-light: stdout is written
 * through an injectable sink and the workbook loader is overridable, which is
 * what makes the `--test` runner able to replay commands and capture results.
 */
import * as path from 'path';
import { ExtractCommand, BookSource, ExtractOp } from './parser/types';
import { Workbook, Sheet, loadWorkbook, workbookFromCsvText } from './engine/workbook';
import {
  extractCell,
  extractRange,
  extractUsedRange,
  extractTitle,
  ExtractionResult,
} from './engine/extract';
import { dispatch, OutputContext } from './output/actions';
import { ExtractConfig, loadConfig } from './config';

export interface RunOptions {
  /** Base directory for resolving relative workbook paths. Defaults to cwd. */
  cwd?: string;
  /** Configuration overrides (e.g. a custom column separator). */
  config?: Partial<ExtractConfig>;
  /** Sink for terminal output. Defaults to `process.stdout.write`. */
  write?: (text: string) => void;
  /** Override how a book source becomes a Workbook (used by tests). */
  loadWorkbook?: (source: BookSource, cwd: string) => Promise<Workbook>;
}

/** Run a full extract command across all of its books. */
export async function runExtract(
  command: ExtractCommand,
  options: RunOptions = {},
): Promise<void> {
  const config = loadConfig(options.config);
  const cwd = options.cwd ?? process.cwd();
  const ctx: OutputContext = {
    config,
    write: options.write ?? ((text) => process.stdout.write(text)),
  };
  const loader = options.loadWorkbook ?? defaultLoader;

  for (const book of command.books) {
    const workbook = await loader(book.source, cwd);

    // Resolve the book's default sheet lazily so a `--range sheet:"Name"` op can
    // target its own sheet on a multi-sheet workbook without forcing --sheet.
    let defaultSheet: Sheet | null = null;
    const getDefaultSheet = (): Sheet => {
      if (!defaultSheet) defaultSheet = workbook.resolveSheet(book.sheet);
      return defaultSheet;
    };

    const results = book.ops.map((op) => {
      const sheet =
        op.type === 'range' && op.sheetName
          ? workbook.resolveSheet(op.sheetName)
          : getDefaultSheet();
      return runOp(op, sheet, config);
    });

    await dispatch(results, book.action, ctx);
  }
}

/** Execute a single extract op against a resolved sheet. */
export function runOp(
  op: ExtractOp,
  sheet: Sheet,
  config: ExtractConfig,
): ExtractionResult {
  const result = runOpInner(op, sheet, config);
  // Carry the op's assume-merge state onto the result for the aligned renderer.
  result.assumeMerge = op.assumeMerge ?? false;
  return result;
}

function runOpInner(
  op: ExtractOp,
  sheet: Sheet,
  config: ExtractConfig,
): ExtractionResult {
  switch (op.type) {
    case 'cell':
      return extractCell(sheet, op.ref);
    case 'range':
      if (op.usedRange) return extractUsedRange(sheet);
      return extractRange(sheet, op.ref ?? '');
    case 'title':
      return extractTitle(sheet, op.title, { headerRow: op.headerRow, config });
  }
}

/** Default workbook resolution from a book source. */
async function defaultLoader(source: BookSource, cwd: string): Promise<Workbook> {
  if (source.kind === 'raw') {
    return workbookFromCsvText(source.text, source.name);
  }
  const resolved = path.isAbsolute(source.path)
    ? source.path
    : path.resolve(cwd, source.path);
  return loadWorkbook(resolved);
}
