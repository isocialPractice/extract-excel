/**
 * Tokenizer: turns a raw argv array into a structured {@link Command}.
 *
 * The grammar is positional with sticky semantics, summarized from the draft:
 *   - The first bare (non-flag) token defaults to `--book` (the workbook).
 *   - Each `-b/--book` starts a NEW book; following extract/output options
 *     attach to that book until the next book.
 *   - `-s/--sheet` is only honored when it immediately follows its book.
 *   - `-a/--action` carries colon-joined targets and `k=v` path arguments.
 *   - `-h/--help` and `--test` are application options handled separately.
 */
import {
  ActionSpec,
  ActionTarget,
  BookJob,
  Command,
  ExtractCommand,
  ExtractOp,
  HelpCommand,
  OutputFormat,
  PageOrientation,
  SheetQuery,
  TestCommand,
  defaultAction,
} from './types';
import { ExtractError } from '../errors';

const KNOWN_FLAGS: Record<string, string> = {
  '-a': 'action',
  '--action': 'action',
  '-b': 'book',
  '--book': 'book',
  '-c': 'cell',
  '--cell': 'cell',
  '-f': 'file',
  '--file': 'file',
  '-h': 'help',
  '--help': 'help',
  '-o': 'orientation',
  '--orientation': 'orientation',
  '-r': 'range',
  '--range': 'range',
  '-s': 'sheet',
  '--sheet': 'sheet',
  '-t': 'title',
  '--title': 'title',
  '-v': 'version',
  '--version': 'version',
  '-x': 'xml',
  '--xml': 'xml',
  '--assume-merge': 'assume-merge',
  '--test': 'test',
};

const ARG_TARGETS: ActionTarget[] = ['file', 'pdf', 'md', 'var'];
const ALL_TARGETS: ActionTarget[] = ['stdout', 'file', 'pdf', 'md', 'var'];
/**
 * Format modifiers that may appear among `--action:` targets. `md` is no longer
 * a format alias here — it now names the markdown-file output target, so only
 * the explicit `markdown` token selects the markdown format.
 */
const FORMATS: Record<string, OutputFormat> = {
  text: 'text',
  table: 'table',
  csv: 'csv',
  markdown: 'markdown',
  xml: 'xml',
};

const ORIENTATIONS: PageOrientation[] = ['landscape', 'portrait'];

interface Flag {
  name: string;
  /**
   * Text after the first `:` or `=` separator, e.g. `file,stdout` for
   * `--action:file,stdout`, or `false` for `--assume-merge=false`.
   */
  suffix?: string;
}

/** Classify a token as a known flag (honoring the `:` and `=` suffix forms). */
function classifyFlag(token: string): Flag | null {
  if (!token.startsWith('-')) return null;
  // Split on whichever of `:` (action targets) or `=` (switch values) comes
  // first; bare value tokens never start with `-`, so this is unambiguous.
  const colon = token.indexOf(':');
  const equals = token.indexOf('=');
  const sep =
    colon === -1 ? equals : equals === -1 ? colon : Math.min(colon, equals);
  const head = sep === -1 ? token : token.slice(0, sep);
  const name = KNOWN_FLAGS[head];
  if (!name) return null;
  return { name, suffix: sep === -1 ? undefined : token.slice(sep + 1) };
}

/**
 * Map a `--file` path to the action target it implies. A trailing `.pdf` or
 * `.md` extension routes the output through the aligned PDF/markdown renderer;
 * anything else (including `name.md.txt`, whose final extension is `.txt`) is a
 * plain text file.
 */
function impliedFileTarget(filePath: string): 'pdf' | 'md' | 'file' {
  const lower = filePath.toLowerCase();
  if (lower.endsWith('.pdf')) return 'pdf';
  if (lower.endsWith('.md')) return 'md';
  return 'file';
}

/** Parse a `switch:bool` suffix: bare/absent => true, else `true`/`false`. */
function parseSwitchBool(suffix: string | undefined, flag: string): boolean {
  if (suffix === undefined || suffix === '') return true;
  const lower = suffix.toLowerCase();
  if (lower === 'true') return true;
  if (lower === 'false') return false;
  throw new ExtractError(
    'INVALID_ARGUMENT',
    `Option "${flag}" expects true or false, got "${suffix}".`,
    `Use ${flag}, ${flag}=true, or ${flag}=false.`,
  );
}

/** True when `name` is a usable (simple) XML element name. */
function isValidXmlName(name: string): boolean {
  return /^[A-Za-z_][\w.-]*$/.test(name);
}

/**
 * Parse the optional `--xml:root,row` suffix into a tag-name override. The first
 * value names the document root, the second the repeating row element; either
 * may be omitted (`--xml:,Record` overrides only the row). Returns `undefined`
 * when no suffix is present, and throws on an invalid XML element name.
 */
function parseXmlOverride(
  suffix: string | undefined,
): { root?: string; row?: string } | undefined {
  if (suffix === undefined || suffix.trim() === '') return undefined;
  const [rawRoot, rawRow] = suffix.split(',');
  const root = rawRoot?.trim() || undefined;
  const row = rawRow?.trim() || undefined;
  for (const name of [root, row]) {
    if (name !== undefined && !isValidXmlName(name)) {
      throw new ExtractError(
        'INVALID_ARGUMENT',
        `Invalid XML tag name "${name}" in --xml.`,
        'Use --xml:root,row with valid XML element names (letters, digits, _, -, .).',
      );
    }
  }
  return root === undefined && row === undefined ? undefined : { root, row };
}

/**
 * Split a quoted argument string (from `--test custom:.. "<args>"`) into argv,
 * honoring single and double quotes. Mirrors basic shell word-splitting.
 */
export function splitArgString(input: string): string[] {
  const args: string[] = [];
  let cur = '';
  let quote: '"' | "'" | null = null;
  let has = false;
  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    if (quote) {
      if (ch === quote) quote = null;
      else cur += ch;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      has = true;
    } else if (ch === ' ' || ch === '\t') {
      if (has || cur.length > 0) {
        args.push(cur);
        cur = '';
        has = false;
      }
    } else {
      cur += ch;
      has = true;
    }
  }
  if (has || cur.length > 0) args.push(cur);
  return args;
}

/** Entry point: parse a full argv array into a Command. */
export function parse(argv: string[]): Command {
  const tokens = [...argv];
  const first = tokens[0];
  const firstFlag = first !== undefined ? classifyFlag(first) : null;

  if (firstFlag?.name === 'help') return parseHelp(tokens);
  if (firstFlag?.name === 'test') return parseTest(tokens);
  if (firstFlag?.name === 'version') return { kind: 'version' };

  return parseExtract(tokens);
}

function parseHelp(tokens: string[]): HelpCommand {
  const topicArg = tokens[1]?.toLowerCase();
  let topic: HelpCommand['topic'] = 'global';
  if (topicArg === 'opt' || topicArg === 'option' || topicArg === 'options') {
    topic = 'opt';
  } else if (topicArg === 'quick' || topicArg === 'quickstart') {
    topic = 'quick';
  }
  return { kind: 'help', topic };
}

function parseTest(tokens: string[]): TestCommand {
  // tokens[0] is `--test`; the spec (if any) is tokens[1].
  const spec = tokens[1];
  if (spec === undefined) {
    return { kind: 'test', mode: 'global', filters: {} };
  }

  const colon = spec.indexOf(':');
  const head = colon === -1 ? spec : spec.slice(0, colon);
  const tail = colon === -1 ? undefined : spec.slice(colon + 1);

  if (head === 'unit') {
    const filters: Record<string, string> = {};
    for (let i = 2; i < tokens.length; i++) {
      const m = /^([A-Za-z]+)=(.+)$/.exec(tokens[i]);
      if (!m) break;
      filters[m[1].toLowerCase()] = m[2];
    }
    return { kind: 'test', mode: 'unit', category: tail, filters };
  }

  if (head === 'custom') {
    return {
      kind: 'test',
      mode: 'custom',
      filters: {},
      customTarget: tail ?? 'single',
      replay: tokens[2],
    };
  }

  // `--test global` or any unrecognized spec falls back to the global suite.
  return { kind: 'test', mode: 'global', filters: {} };
}

/**
 * Per-book `--assume-merge` bookkeeping. The option is sticky and positional,
 * but with a convenience rule from the spec: when it appears at most once in a
 * book it applies to *every* extraction in that book regardless of position;
 * when it appears more than once each op takes the value active when it parsed.
 */
interface MergeState {
  /** Number of `--assume-merge` occurrences seen in this book. */
  count: number;
  /** The currently active (sticky) value, applied to ops as they are parsed. */
  current: boolean;
  /** The value of the single occurrence, used by the "applies anywhere" rule. */
  single: boolean;
}

function parseExtract(tokens: string[]): ExtractCommand {
  const books: BookJob[] = [];
  let current: BookJob | null = null;
  /** True only on the token immediately after a book is opened. */
  let sheetAllowed = false;
  /** Tracks whether the user has customized output (so default stdout drops). */
  const customized = new WeakSet<BookJob>();
  /** Per-book assume-merge state (see {@link MergeState}). */
  const mergeState = new WeakMap<BookJob, MergeState>();

  const openBook = (path: string): void => {
    current = { source: { kind: 'file', path }, ops: [], action: defaultAction() };
    books.push(current);
    mergeState.set(current, { count: 0, current: false, single: false });
    sheetAllowed = true;
  };

  /** Push an op onto a book, tagging it with the active assume-merge state. */
  const pushOp = (book: BookJob, op: ExtractOp): void => {
    op.assumeMerge = mergeState.get(book)?.current ?? false;
    book.ops.push(op);
  };

  const requireBook = (what: string): BookJob => {
    if (!current) {
      throw new ExtractError(
        'NO_WORKBOOK',
        `Cannot apply ${what} before a workbook is specified.`,
        'Provide a workbook first, e.g. `extract-excel file.xlsx --cell A1`.',
      );
    }
    return current;
  };

  const applyAction = (book: BookJob, spec: ActionSpec): void => {
    if (!customized.has(book)) {
      // First explicit output option replaces the implicit stdout default.
      book.action = { targets: [] };
      customized.add(book);
    }
    for (const t of spec.targets) {
      if (!book.action.targets.includes(t)) book.action.targets.push(t);
    }
    if (spec.format !== undefined) book.action.format = spec.format;
    if (spec.file !== undefined) book.action.file = spec.file;
    if (spec.pdf !== undefined) book.action.pdf = spec.pdf;
    if (spec.md !== undefined) book.action.md = spec.md;
    if (spec.var !== undefined) book.action.var = spec.var;
    if (spec.orientation !== undefined) book.action.orientation = spec.orientation;
    if (spec.xml !== undefined) book.action.xml = spec.xml;
  };

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    const flag = classifyFlag(token);

    // A non-flag token is only meaningful as the very first (default book).
    if (!flag) {
      if (i === 0) {
        openBook(token);
        continue;
      }
      throw new ExtractError(
        'INVALID_ARGUMENT',
        `Unexpected argument "${token}".`,
        'Workbooks after the first must be introduced with -b/--book.',
      );
    }

    const next = (): string => {
      const v = tokens[i + 1];
      if (v === undefined) {
        throw new ExtractError('INVALID_ARGUMENT', `Option "${token}" expects a value.`);
      }
      i++;
      return v;
    };

    switch (flag.name) {
      case 'book': {
        openBook(next());
        break;
      }
      case 'sheet': {
        // Qualifier form: --sheet:length, --sheet:list, --sheet:info —
        // workbook introspection queries that don't select a sheet for extraction.
        if (flag.suffix === 'length' || flag.suffix === 'list' || flag.suffix === 'info') {
          requireBook(`--sheet:${flag.suffix}`).sheetQuery = flag.suffix as SheetQuery;
          sheetAllowed = false;
          break;
        }
        // Normal sheet name selection (must immediately follow the book token).
        const name = next();
        // Derive the open book from `books` to sidestep closure-assignment CFA.
        const book = books.length > 0 ? books[books.length - 1] : null;
        if (sheetAllowed && book) book.sheet = name;
        // Whether honored or ignored, the sheet window closes after this.
        sheetAllowed = false;
        break;
      }
      case 'cell': {
        pushOp(requireBook('--cell'), { type: 'cell', ref: next() });
        sheetAllowed = false;
        break;
      }
      case 'range': {
        const book = requireBook('--range');
        const value = next();
        if (value === 'sheet') {
          // `--range sheet` => the (single) sheet's whole used range.
          pushOp(book, { type: 'range', usedRange: true });
        } else if (value.startsWith('sheet:')) {
          // `--range sheet:"Name"` => a named sheet's whole used range.
          pushOp(book, { type: 'range', usedRange: true, sheetName: value.slice(6) });
        } else {
          pushOp(book, { type: 'range', ref: value });
        }
        sheetAllowed = false;
        break;
      }
      case 'title': {
        const book = requireBook('--title');
        let headerRow: number | undefined;
        // Optional `row=N` precedes the title (e.g. `--title row=2 "March"`).
        const rowMatch = /^row=(\d+)$/.exec(tokens[i + 1] ?? '');
        if (rowMatch) {
          headerRow = parseInt(rowMatch[1], 10);
          i++;
        }
        pushOp(book, { type: 'title', title: next(), headerRow });
        sheetAllowed = false;
        break;
      }
      case 'file': {
        const book = requireBook('--file');
        const value = next();
        // A `.pdf`/`.md` path implies the aligned PDF/markdown target.
        const target = impliedFileTarget(value);
        applyAction(book, { targets: [target], [target]: value });
        sheetAllowed = false;
        break;
      }
      case 'action': {
        const book = requireBook('--action');
        const { spec, consumed } = parseActionArgs(flag.suffix, tokens, i + 1);
        i += consumed;
        applyAction(book, spec);
        sheetAllowed = false;
        break;
      }
      case 'assume-merge': {
        const book = requireBook('--assume-merge');
        const value = parseSwitchBool(flag.suffix, '--assume-merge');
        const state = mergeState.get(book);
        if (state) {
          state.count++;
          state.current = value;
          state.single = value;
        }
        sheetAllowed = false;
        break;
      }
      case 'orientation': {
        const book = requireBook('--orientation');
        const value = next().toLowerCase();
        if (!ORIENTATIONS.includes(value as PageOrientation)) {
          throw new ExtractError(
            'INVALID_ARGUMENT',
            `Option "${token}" expects landscape or portrait, got "${value}".`,
            'Page orientation only applies to PDF output.',
          );
        }
        // Route through applyAction so the value survives the implicit-stdout
        // reset that the first explicit output option performs.
        applyAction(book, { targets: [], orientation: value as PageOrientation });
        sheetAllowed = false;
        break;
      }
      case 'xml': {
        // `-x/--xml` is shorthand for the `--action:xml` format modifier. It
        // selects the XML format but adds no target of its own (routed with an
        // empty target list, like --orientation), so a following --file/--action
        // chooses the destination and a bare --xml falls back to stdout via the
        // dispatch default. An optional `--xml:root,row` suffix overrides the
        // detected container tag names.
        const override = parseXmlOverride(flag.suffix);
        applyAction(requireBook('--xml'), {
          targets: [],
          format: 'xml',
          ...(override ? { xml: override } : {}),
        });
        sheetAllowed = false;
        break;
      }
      default:
        // help/test/version handled before parseExtract; nothing reaches here.
        break;
    }
  }

  // Apply the "used at most once => applies to every op in the book" rule.
  for (const book of books) {
    const state = mergeState.get(book);
    if (state && state.count <= 1) {
      const value = state.count === 1 ? state.single : false;
      for (const op of book.ops) op.assumeMerge = value;
    }
  }

  // `--xml`/`--action:xml` map a whole sheet by default. With no explicit
  // extraction op, fall back to the sheet's used range so the header row can
  // supply the XML field names (`extract-excel file.xlsx --xml`).
  for (const book of books) {
    if (book.action.format === 'xml' && book.ops.length === 0) {
      book.ops.push({ type: 'range', usedRange: true, assumeMerge: false });
    }
  }

  if (books.length === 0) {
    throw new ExtractError(
      'NO_WORKBOOK',
      'No workbook was specified.',
      'Run `extract-excel --help` for usage.',
    );
  }
  return { kind: 'extract', books };
}

/**
 * Parse the targets + arguments for one `--action` occurrence.
 *
 * - `suffix` holds the colon-joined targets (`file,stdout`); absent => stdout.
 * - Targets needing a value are file/pdf/var. With exactly one such target the
 *   value is the next positional token. With several, named `k=v` tokens are
 *   consumed (`var=_n file=path pdf=path`) and each arg target must be filled.
 *
 * Returns the parsed spec plus how many tokens after the flag were consumed,
 * so the caller can advance its cursor.
 */
function parseActionArgs(
  suffix: string | undefined,
  tokens: string[],
  start: number,
): { spec: ActionSpec; consumed: number } {
  const tokensInSuffix = (suffix ? suffix.split(',') : [])
    .map((t) => t.trim())
    .filter(Boolean);

  const targets: ActionTarget[] = [];
  let format: OutputFormat | undefined;
  for (const t of tokensInSuffix) {
    const lower = t.toLowerCase();
    if (ALL_TARGETS.includes(t as ActionTarget)) {
      targets.push(t as ActionTarget);
    } else if (FORMATS[lower]) {
      format = FORMATS[lower];
    } else {
      throw new ExtractError(
        'MALFORMED_ACTION',
        `Unknown action target or format "${t}".`,
        'Targets: stdout, file, pdf, md, var. Formats: text, table, csv, markdown, xml.',
      );
    }
  }

  // A format-only action (e.g. `--action:xml`, `--action:table`) names no
  // destination. If a path follows it (`--action:xml file.xml`) route the output
  // to that file; otherwise default to the terminal. Bare tokens after the first
  // argument are otherwise rejected, so this only gives meaning to a form that
  // previously errored — it never changes an already-valid command.
  if (targets.length === 0) {
    const following = tokens[start];
    if (format && following !== undefined && !following.startsWith('-')) {
      targets.push('file');
    } else {
      targets.push('stdout');
    }
  }

  const spec: ActionSpec = format ? { targets, format } : { targets };
  const argTargets = targets.filter((t) => ARG_TARGETS.includes(t));
  let consumed = 0;

  if (argTargets.length === 1) {
    const value = tokens[start];
    if (value === undefined) {
      throw new ExtractError(
        'MALFORMED_ACTION',
        `Action target "${argTargets[0]}" requires a value.`,
      );
    }
    setTarget(spec, argTargets[0], value);
    consumed = 1;
  } else if (argTargets.length > 1) {
    const seen = new Set<ActionTarget>();
    for (let idx = start; idx < tokens.length; idx++) {
      const m = /^(file|pdf|md|var)=(.+)$/.exec(tokens[idx]);
      if (!m) break;
      setTarget(spec, m[1] as ActionTarget, m[2]);
      seen.add(m[1] as ActionTarget);
      consumed++;
    }
    const missing = argTargets.filter((t) => !seen.has(t));
    if (missing.length > 0) {
      throw new ExtractError(
        'MALFORMED_ACTION',
        `Action targets ${missing.join(', ')} require named arguments (e.g. ${missing[0]}=value).`,
        'When combining multiple file/pdf/var targets, pass each as name=value.',
      );
    }
  }

  return { spec, consumed };
}

/** Assign a value to the matching field on an ActionSpec. */
function setTarget(spec: ActionSpec, target: ActionTarget, value: string): void {
  if (target === 'file') spec.file = value;
  else if (target === 'pdf') spec.pdf = value;
  else if (target === 'md') spec.md = value;
  else if (target === 'var') spec.var = value;
}
