/**
 * The structured command model produced by the tokenizer and consumed by the
 * runner. The CLI grammar in the draft doc is bespoke (sticky `--sheet`,
 * repeating books, `--action:targets` with colon-joined targets and `k=v`
 * arguments), so it is parsed into this explicit shape rather than via a
 * generic flag library.
 */

/** One output destination for a book's aggregated extraction text. */
export type ActionTarget = 'stdout' | 'file' | 'pdf' | 'var';

/** How extracted data is rendered to a destination. */
export type OutputFormat = 'text' | 'table' | 'csv' | 'markdown';

export interface ActionSpec {
  /** Where the book's output is directed. Defaults to `['stdout']`. */
  targets: ActionTarget[];
  /**
   * Explicit render format (from an `--action:` modifier such as `table`).
   * When absent the format is inferred per target (e.g. `.csv` => csv,
   * `pdf` => markdown).
   */
  format?: OutputFormat;
  /** Filesystem path for the `file` target. */
  file?: string;
  /** Filesystem path for the `pdf` target. */
  pdf?: string;
  /** Variable name for the `var` target (leading `_` allowed). */
  var?: string;
}

/**
 * A single extraction request within a book.
 *
 * The `range` op also covers the `--range sheet` / `--range sheet:"Name"`
 * presets, which extract a whole sheet's used range (its data bounding box).
 */
export type ExtractOp =
  | { type: 'cell'; ref: string }
  | { type: 'range'; ref?: string; usedRange?: boolean; sheetName?: string }
  | { type: 'title'; title: string; headerRow?: number };

/** Where a book's data comes from. */
export type BookSource =
  | { kind: 'file'; path: string }
  | { kind: 'raw'; text: string; name?: string };

/** A workbook plus everything to extract from it and where output should go. */
export interface BookJob {
  source: BookSource;
  /** Selected sheet name (must have immediately followed the book token). */
  sheet?: string;
  ops: ExtractOp[];
  /** Output routing; merged from `--file` and `--action` for this book. */
  action: ActionSpec;
}

/** `--help` request. */
export interface HelpCommand {
  kind: 'help';
  /** `global` (full doc), or a specific topic such as `opt`/`quick`. */
  topic: 'global' | 'opt' | 'quick';
}

/** `--test` request. */
export interface TestCommand {
  kind: 'test';
  mode: 'global' | 'unit' | 'custom';
  /** For `unit:extract` etc. — the task category to filter by. */
  category?: string;
  /** Fine-grained unit filters: `type=multi`, `opt=range`. */
  filters: Record<string, string>;
  /** For `custom:<target>` — `single`, `multi`, a file path, or raw CSV. */
  customTarget?: string;
  /** The quoted argument string replayed through the engine for custom tests. */
  replay?: string;
}

/** An extraction run across one or more books. */
export interface ExtractCommand {
  kind: 'extract';
  books: BookJob[];
}

export type Command = HelpCommand | TestCommand | ExtractCommand;

/** Create the default action (terminal output). */
export function defaultAction(): ActionSpec {
  return { targets: ['stdout'] };
}
