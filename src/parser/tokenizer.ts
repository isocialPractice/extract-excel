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
  HelpCommand,
  OutputFormat,
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
  '-r': 'range',
  '--range': 'range',
  '-s': 'sheet',
  '--sheet': 'sheet',
  '-t': 'title',
  '--title': 'title',
  '--test': 'test',
};

const ARG_TARGETS: ActionTarget[] = ['file', 'pdf', 'var'];
const ALL_TARGETS: ActionTarget[] = ['stdout', 'file', 'pdf', 'var'];
/** Format modifiers that may appear among `--action:` targets. */
const FORMATS: Record<string, OutputFormat> = {
  text: 'text',
  table: 'table',
  csv: 'csv',
  markdown: 'markdown',
  md: 'markdown',
};

interface Flag {
  name: string;
  /** Text after the first colon, e.g. `file,stdout` for `--action:file,stdout`. */
  suffix?: string;
}

/** Classify a token as a known flag (honoring the `:` suffix form). */
function classifyFlag(token: string): Flag | null {
  if (!token.startsWith('-')) return null;
  const colon = token.indexOf(':');
  const head = colon === -1 ? token : token.slice(0, colon);
  const name = KNOWN_FLAGS[head];
  if (!name) return null;
  return { name, suffix: colon === -1 ? undefined : token.slice(colon + 1) };
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

function parseExtract(tokens: string[]): ExtractCommand {
  const books: BookJob[] = [];
  let current: BookJob | null = null;
  /** True only on the token immediately after a book is opened. */
  let sheetAllowed = false;
  /** Tracks whether the user has customized output (so default stdout drops). */
  const customized = new WeakSet<BookJob>();

  const openBook = (path: string): void => {
    current = { source: { kind: 'file', path }, ops: [], action: defaultAction() };
    books.push(current);
    sheetAllowed = true;
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
    if (spec.var !== undefined) book.action.var = spec.var;
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
        const name = next();
        // Derive the open book from `books` to sidestep closure-assignment CFA.
        const book = books.length > 0 ? books[books.length - 1] : null;
        if (sheetAllowed && book) book.sheet = name;
        // Whether honored or ignored, the sheet window closes after this.
        sheetAllowed = false;
        break;
      }
      case 'cell': {
        requireBook('--cell').ops.push({ type: 'cell', ref: next() });
        sheetAllowed = false;
        break;
      }
      case 'range': {
        const book = requireBook('--range');
        const value = next();
        if (value === 'sheet') {
          // `--range sheet` => the (single) sheet's whole used range.
          book.ops.push({ type: 'range', usedRange: true });
        } else if (value.startsWith('sheet:')) {
          // `--range sheet:"Name"` => a named sheet's whole used range.
          book.ops.push({ type: 'range', usedRange: true, sheetName: value.slice(6) });
        } else {
          book.ops.push({ type: 'range', ref: value });
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
        book.ops.push({ type: 'title', title: next(), headerRow });
        sheetAllowed = false;
        break;
      }
      case 'file': {
        const book = requireBook('--file');
        applyAction(book, { targets: ['file'], file: next() });
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
      default:
        // help/test handled before parseExtract; nothing else should reach here.
        break;
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
        'Targets: stdout, file, pdf, var. Formats: text, table, csv, markdown.',
      );
    }
  }

  // A format-only action (e.g. `--action:table`) still defaults to the terminal.
  if (targets.length === 0) targets.push('stdout');

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
      const m = /^(file|pdf|var)=(.+)$/.exec(tokens[idx]);
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
  else if (target === 'var') spec.var = value;
}
