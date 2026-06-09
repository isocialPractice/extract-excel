/**
 * Execution layer: run a parsed {@link ExtractCommand} end-to-end.
 *
 * This module is the bridge between the tokenizer's data model and the engine +
 * output layers. It is intentionally side-effect-light: stdout is written
 * through an injectable sink and the workbook loader is overridable, which is
 * what makes the `--test` runner able to replay commands and capture results.
 */
import * as path from 'path';
import { ExtractCommand, BookSource, ExtractOp, ActionSpec, XmlMapping, SheetQuery } from './parser/types';
import {
  Workbook,
  Sheet,
  loadWorkbook,
  workbookFromCsvText,
  readXmlMapping,
} from './engine/workbook';
import {
  extractCell,
  extractRange,
  extractUsedRange,
  extractTitle,
  ExtractionResult,
} from './engine/extract';
import { dispatch, dispatchText, OutputContext } from './output/actions';
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

    // Handle sheet queries (--sheet:length / :list / :info) before extraction.
    if (book.sheetQuery) {
      const targets = book.action.targets;
      const listFormat = targets.includes('md') || targets.includes('pdf');
      const text = renderSheetQuery(workbook, book.sheetQuery, listFormat);
      await dispatchText(text, book.action, ctx);
      continue;
    }

    // For XML output, resolve the container tags: a `--xml:root,row` override
    // layered over the workbook's embedded Excel XML map (when it has one). The
    // map is read straight from the source file, so it only applies to file
    // sources (not raw CSV) and only when XML output was requested.
    const sourcePath =
      book.source.kind === 'file'
        ? path.resolve(cwd, book.source.path)
        : undefined;
    const xmlMapping =
      book.action.format === 'xml'
        ? await resolveXmlMapping(book.action, sourcePath)
        : undefined;

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
      const result = runOp(op, sheet, config);
      if (xmlMapping) result.xmlMapping = xmlMapping;
      return result;
    });

    await dispatch(results, book.action, ctx);
  }
}

/**
 * Resolve the effective XML container mapping for a book: the user's
 * `--xml:root,row` override takes precedence, falling back to the workbook's
 * embedded Excel XML map. Returns `undefined` when neither supplies a name, so
 * the renderer uses its sheet-name/`row` defaults.
 */
async function resolveXmlMapping(
  action: ActionSpec,
  sourcePath: string | undefined,
): Promise<XmlMapping | undefined> {
  const override = action.xml;
  const detected = sourcePath ? await readXmlMapping(sourcePath) : undefined;
  if (!override && !detected) return undefined;
  return {
    root: override?.root ?? detected?.root,
    row: override?.row ?? detected?.row,
    // Emit the schema-instance namespace whenever a map or override is in play,
    // mirroring Excel's own XML export. A detected map supplies its own set.
    namespaces: detected?.namespaces ?? {
      'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    },
  };
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
  // Record the source sheet so structured renderers (xml) can name the output.
  result.sheetName = sheet.name;
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

/**
 * Render workbook sheet metadata as plain text for a `--sheet` query.
 *
 * - `length` : the sheet count as a bare number.
 * - `list`   : one sheet name per line; with `listFormat` each line is prefixed
 *              with `- ` so it reads as a markdown bullet list.
 * - `info`   : `length: N` header followed by the bullet list of names.
 */
function renderSheetQuery(workbook: Workbook, query: SheetQuery, listFormat: boolean): string {
  const names = workbook.sheetNames;
  const count = workbook.sheetCount;
  switch (query) {
    case 'length':
      return String(count);
    case 'list':
      return listFormat
        ? names.map((n) => `- ${n}`).join('\n')
        : names.join('\n');
    case 'info': {
      const lines = [`length: ${count}`, ...names.map((n) => `- ${n}`)];
      return lines.join('\n');
    }
  }
}
