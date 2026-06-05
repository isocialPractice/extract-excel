/**
 * The structured command model produced by the tokenizer and consumed by the
 * runner. The CLI grammar in the draft doc is bespoke (sticky `--sheet`,
 * repeating books, `--action:targets` with colon-joined targets and `k=v`
 * arguments), so it is parsed into this explicit shape rather than via a
 * generic flag library.
 */

/**
 * One output destination for a book's aggregated extraction text.
 *
 * `pdf` and `md` share the same two-pass aligned-table renderer; they differ
 * only in their container (a PDF document vs. a `.md` text file).
 */
export type ActionTarget = 'stdout' | 'file' | 'pdf' | 'md' | 'var';

/** How extracted data is rendered to a destination. */
export type OutputFormat = 'text' | 'table' | 'csv' | 'markdown' | 'xml';

/** Page orientation for the `pdf` target (`-o/--orientation`). */
export type PageOrientation = 'landscape' | 'portrait';

/**
 * Element names for the `xml` format's container tags.
 *
 * In Excel-mapped workbooks these are detected from the embedded XML map
 * (`xl/xmlMaps.xml`): `root` is the document root (e.g. `cardholders`) and `row`
 * is the repeating record element (e.g. `CardHolder`). The user can override
 * either via `--xml:root,row`. `namespaces` holds root attributes such as
 * `xmlns:xsi` so the output mirrors Excel's own XML export.
 */
export interface XmlMapping {
  /** Document root element name. Falls back to the camel-cased sheet name. */
  root?: string;
  /** Repeating record element name. Falls back to `row`. */
  row?: string;
  /** Namespace (and other) attributes to emit on the root element. */
  namespaces?: Record<string, string>;
}

export interface ActionSpec {
  /** Where the book's output is directed. Defaults to `['stdout']`. */
  targets: ActionTarget[];
  /**
   * Explicit render format (from an `--action:` modifier such as `table`).
   * When absent the format is inferred per target (e.g. `.csv` => csv,
   * `pdf`/`md` => the aligned table).
   */
  format?: OutputFormat;
  /** Filesystem path for the `file` target. */
  file?: string;
  /** Filesystem path for the `pdf` target. */
  pdf?: string;
  /** Filesystem path for the `md` target. */
  md?: string;
  /** Variable name for the `var` target (leading `_` allowed). */
  var?: string;
  /**
   * Page orientation for the `pdf` target (`-o/--orientation`). Ignored unless
   * the book has a PDF output (set or implied via a `.pdf` file).
   */
  orientation?: PageOrientation;
  /**
   * User-supplied override of the `xml` container tags (`--xml:root,row`). Only
   * the names the user provided are set; anything absent falls back to the
   * detected XML map, then to the sheet-name/`row` defaults.
   */
  xml?: { root?: string; row?: string };
}

/**
 * A single extraction request within a book.
 *
 * The `range` op also covers the `--range sheet` / `--range sheet:"Name"`
 * presets, which extract a whole sheet's used range (its data bounding box).
 *
 * `assumeMerge` records the `--assume-merge` state that was active for this op
 * (see the tokenizer for the sticky/positional resolution). It only affects the
 * aligned `pdf`/`md` renderer, which collapses adjacent duplicate text cells
 * (merge artifacts) when it is set.
 */
export type ExtractOp =
  | { type: 'cell'; ref: string; assumeMerge?: boolean }
  | {
      type: 'range';
      ref?: string;
      usedRange?: boolean;
      sheetName?: string;
      assumeMerge?: boolean;
    }
  | { type: 'title'; title: string; headerRow?: number; assumeMerge?: boolean };

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

/** `--version` request. */
export interface VersionCommand {
  kind: 'version';
}

/** An extraction run across one or more books. */
export interface ExtractCommand {
  kind: 'extract';
  books: BookJob[];
}

export type Command =
  | HelpCommand
  | TestCommand
  | VersionCommand
  | ExtractCommand;

/** Create the default action (terminal output). */
export function defaultAction(): ActionSpec {
  return { targets: ['stdout'] };
}
