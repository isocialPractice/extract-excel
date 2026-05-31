/**
 * Output dispatch: render a book's results and send them to each target.
 *
 * Format resolution per target:
 *   - pdf    : the source workbook is exported to a styled PDF table via a
 *              Python 3 script (`scripts/xlsx_to_pdf.py`) using openpyxl and
 *              reportlab. Python 3 must be installed, along with the packages
 *              `openpyxl` and `reportlab` (installed automatically by
 *              `npm install`). Raw CSV sources are not supported for pdf.
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
import { execFile } from 'child_process';
import { promisify } from 'util';
import { ActionSpec } from '../parser/types';
import { ExtractConfig } from '../config';
import { ExtractError } from '../errors';
import { ExtractionResult } from '../engine/extract';
import { render, formatFromPath } from './render';

const execFileAsync = promisify(execFile);

export interface OutputContext {
  config: ExtractConfig;
  /** Sink for stdout (overridable so the test runner can capture output). */
  write: (text: string) => void;
  /**
   * Resolved absolute path of the source workbook file. Populated only for
   * file-based sources; `undefined` for raw CSV input. Required for the `pdf`
   * target (Python export).
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
        await convertWithPython(ctx.sourcePath, requirePath(action.pdf, 'pdf'));
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

/** True when the OS could not locate the given executable. */
function isCommandNotFound(err: unknown): boolean {
  const e = err as { code?: number | string; message?: string };
  return e.code === 'ENOENT' || e.code === 127 || e.code === 9009;
}

/**
 * Return the absolute Python 3 path written by the postinstall script
 * (`scripts/setup-pdf.js`), or `null` if the cache is missing or stale.
 */
function cachedPythonPath(): string | null {
  try {
    // Compiled layout: dist/output/actions.js → ../../.python-path
    const cachePath = path.join(__dirname, '..', '..', '.python-path');
    const p = fs.readFileSync(cachePath, 'utf8').trim();
    return p && fs.existsSync(p) ? p : null;
  } catch {
    return null;
  }
}

/**
 * Absolute path to the bundled Python conversion script.
 * Compiled layout: dist/output/actions.js → ../../scripts/xlsx_to_pdf.py
 */
function pdfScriptPath(): string {
  return path.join(__dirname, '..', '..', 'scripts', 'xlsx_to_pdf.py');
}

/**
 * Export `sourcePath` to a styled PDF at `outputPath` using the bundled
 * Python script (openpyxl + reportlab).
 *
 * Resolution order for the Python 3 executable:
 *   1. Path cached by `npm install` (postinstall script)
 *   2. `python3` via system PATH
 *   3. `python` via system PATH (Python 3 only — verified by exit code)
 *
 * Throws `PYTHON_NOT_FOUND` when no usable Python 3 is available, or when
 * the required pip packages are missing (script exit code 3).
 */
async function convertWithPython(
  sourcePath: string,
  outputPath: string,
): Promise<void> {
  const absSource = path.resolve(sourcePath);
  const absOutput = path.resolve(outputPath);
  const script    = pdfScriptPath();

  fs.mkdirSync(path.dirname(absOutput), { recursive: true });

  const cached     = cachedPythonPath();
  const candidates = [...(cached ? [cached] : []), 'python3', 'python'];

  for (const bin of candidates) {
    try {
      await execFileAsync(bin, [script, absSource, absOutput], {
        env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
      });
      return; // success
    } catch (err) {
      if (isCommandNotFound(err)) continue;

      const e = err as { code?: number; stderr?: string; message?: string };

      // Exit code 3: the Python script could not import openpyxl or reportlab.
      if (e.code === 3) {
        throw new ExtractError(
          'PYTHON_NOT_FOUND',
          'Required Python packages (openpyxl, reportlab) are not installed.',
          [
            'Run: python3 -m pip install openpyxl reportlab',
            'Or re-run `npm install` to install them automatically.',
          ].join('\n  '),
        );
      }

      // Any other non-zero exit: surface stderr as the message.
      throw new ExtractError(
        'MALFORMED_ACTION',
        `PDF generation failed: ${e.stderr?.trim() || e.message || String(err)}`,
      );
    }
  }

  throw new ExtractError(
    'PYTHON_NOT_FOUND',
    'Python 3 was not found on this system.',
    [
      'Install Python 3: https://www.python.org/downloads/',
      'Then run `npm install` to cache its path, or add python3 to your PATH.',
      'After installing Python, also run: python3 -m pip install openpyxl reportlab',
    ].join('\n  '),
  );
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
