/**
 * Output dispatch: render a book's results and send them to each target.
 *
 * Format resolution per target:
 *   - pdf    : the source workbook is exported directly to PDF via LibreOffice
 *              headless (`libreoffice --headless --convert-to pdf`). Requires
 *              LibreOffice to be installed with `libreoffice` or `soffice` on
 *              the PATH. Raw CSV sources are not supported for this target.
 *   - file   : the action's explicit format, else inferred from the extension
 *              (`.csv` => csv, `.md` => markdown), else text.
 *   - stdout : the action's explicit format, else text.
 *   - var    : always text, collapsed to a single sourceable assignment. A
 *              child process cannot set a variable in its parent shell, so the
 *              snippet is emitted for the user to `eval` / `for /f` (see the
 *              capture helpers under `scripts/`).
 */
import * as fs from 'fs';
import * as path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import { ActionSpec } from '../parser/types';
import { ExtractConfig } from '../config';
import { ExtractError } from '../errors';
import { ExtractionResult } from '../engine/extract';
import { render, formatFromPath } from './render';

const execAsync = promisify(exec);

export interface OutputContext {
  config: ExtractConfig;
  /** Sink for stdout (overridable so the test runner can capture output). */
  write: (text: string) => void;
  /**
   * Resolved absolute path of the source workbook file. Populated only for
   * file-based sources; `undefined` for raw CSV input. Required for the `pdf`
   * target (LibreOffice export).
   */
  sourcePath?: string;
}

/** Dispatch one book's results to all of its targets. */
export async function dispatch(
  results: ExtractionResult[],
  action: ActionSpec,
  ctx: OutputContext,
): Promise<void> {
  const targets = action.targets.length > 0 ? action.targets : ['stdout'];
  for (const target of targets) {
    switch (target) {
      case 'stdout':
        ctx.write(render(results, action.format ?? 'text', ctx.config) + '\n');
        break;
      case 'file': {
        const file = requirePath(action.file, 'file');
        const format = action.format ?? formatFromPath(file) ?? 'text';
        writeTextFile(file, render(results, format, ctx.config));
        break;
      }
      case 'pdf': {
        if (!ctx.sourcePath) {
          throw new ExtractError(
            'MALFORMED_ACTION',
            'The pdf target requires a file-based workbook source; raw CSV input cannot be exported as PDF.',
          );
        }
        await convertWithLibreOffice(ctx.sourcePath, requirePath(action.pdf, 'pdf'));
        break;
      }
      case 'var': {
        const text = render(results, 'text', ctx.config);
        ctx.write(renderVarExport(requireName(action.var), text, ctx.config) + '\n');
        break;
      }
    }
  }
}

function requirePath(value: string | undefined, target: string): string {
  if (!value) {
    throw new ExtractError('MALFORMED_ACTION', `The "${target}" action has no path.`);
  }
  return value;
}

function requireName(value: string | undefined): string {
  if (!value) {
    throw new ExtractError('MALFORMED_ACTION', 'The "var" action has no variable name.');
  }
  return value;
}

/** Ensure the parent directory exists, then write the text file. */
function writeTextFile(filePath: string, text: string): void {
  const dir = path.dirname(path.resolve(filePath));
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(filePath, text + '\n', 'utf8');
}

/** True when the shell or OS could not locate the given command. */
function isCommandNotFound(err: unknown): boolean {
  const e = err as { code?: number | string; stderr?: string; message?: string };
  // Numeric exit codes: 127 = POSIX "command not found", 9009 = Windows cmd.exe
  if (e.code === 'ENOENT' || e.code === 127 || e.code === 9009) return true;
  // Windows puts the human-readable "not recognized" text in stderr/message
  // rather than relying on a stable exit code across all cmd.exe versions.
  const text = `${e.stderr ?? ''} ${e.message ?? ''}`;
  return /is not recognized as an internal or external command/i.test(text) ||
    /command not found/i.test(text);
}

/** Well-known default installation paths for LibreOffice per platform. */
const LIBREOFFICE_COMMON_PATHS: readonly string[] = (() => {
  switch (process.platform) {
    case 'win32':
      return [
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
      ];
    case 'darwin':
      return ['/Applications/LibreOffice.app/Contents/MacOS/soffice'];
    default:
      return [
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice',
      ];
  }
})();

/**
 * Return the absolute LibreOffice path written by the postinstall script
 * (`scripts/find-libreoffice.js`), or `null` if the cache is missing/stale.
 */
function cachedLibreOfficePath(): string | null {
  try {
    // Compiled layout: dist/output/actions.js → ../../.libreoffice-path
    const cachePath = path.join(__dirname, '..', '..', '.libreoffice-path');
    const p = fs.readFileSync(cachePath, 'utf8').trim();
    return p && fs.existsSync(p) ? p : null;
  } catch {
    return null;
  }
}

/**
 * Attempt each LibreOffice candidate in order, stopping on first success.
 *
 * Resolution order:
 *   1. Path cached by `npm install` (postinstall script)
 *   2. `libreoffice` / `soffice` via system PATH
 *   3. Common absolute installation paths (runtime fallback)
 *
 * Throws `LIBREOFFICE_NOT_FOUND` when none of the candidates work.
 */
async function spawnLibreOffice(sourcePath: string, outDir: string): Promise<void> {
  const cached   = cachedLibreOfficePath();
  const absolute = LIBREOFFICE_COMMON_PATHS.filter(
    (p) => fs.existsSync(p) && p !== cached,
  );
  const candidates = [
    ...(cached ? [cached] : []),
    'libreoffice',
    'soffice',
    ...absolute,
  ];

  for (const bin of candidates) {
    // Quote executables whose path contains spaces (common on Windows).
    const quoted = /\s/.test(bin) ? `"${bin}"` : bin;
    try {
      await execAsync(
        `${quoted} --headless --convert-to pdf --outdir "${outDir}" "${sourcePath}"`,
      );
      return;
    } catch (err) {
      if (isCommandNotFound(err)) continue;
      throw err;
    }
  }

  throw new ExtractError(
    'LIBREOFFICE_NOT_FOUND',
    'LibreOffice was not found on this system.',
    [
      'Install LibreOffice: https://www.libreoffice.org/download/libreoffice-still/',
      'Then run `npm install` to cache its path, or add it to your PATH.',
      `Default locations: ${LIBREOFFICE_COMMON_PATHS.join(', ')}`,
    ].join('\n  '),
  );
}

/**
 * Export `sourcePath` to PDF at `outputPath` using LibreOffice headless.
 *
 * LibreOffice always places its output as `<source-stem>.pdf` inside the
 * output directory. If that name differs from the requested `outputPath` the
 * file is renamed to match.
 */
async function convertWithLibreOffice(
  sourcePath: string,
  outputPath: string,
): Promise<void> {
  const absSource = path.resolve(sourcePath);
  const absOutput = path.resolve(outputPath);
  const outDir = path.dirname(absOutput);

  fs.mkdirSync(outDir, { recursive: true });
  await spawnLibreOffice(absSource, outDir);

  // Rename if LibreOffice's implicit <stem>.pdf differs from the requested path.
  const stem = path.basename(absSource, path.extname(absSource));
  const libreOut = path.join(outDir, `${stem}.pdf`);
  if (path.resolve(libreOut) !== absOutput) {
    fs.renameSync(libreOut, absOutput);
  }
}

/**
 * Produce an OS-appropriate variable assignment snippet. Multi-line values are
 * collapsed onto one line (joined with `; `) so they remain assignable.
 */
export function renderVarExport(
  name: string,
  text: string,
  config: ExtractConfig,
): string {
  const value = text.replace(/\r?\n/g, '; ').trim();
  const style =
    config.varExport === 'auto'
      ? process.platform === 'win32'
        ? 'cmd'
        : 'posix'
      : config.varExport;

  switch (style) {
    case 'cmd':
      return `set "${name}=${value}"`;
    case 'powershell':
      return `$env:${name} = "${value}"`;
    case 'posix':
    default:
      return `export ${name}="${value}"`;
  }
}
