/**
 * Output dispatch: render a book's results and send them to each target.
 *
 * Format resolution per target:
 *   - pdf    : the two-pass aligned table rendered into a PDF (honors
 *              `-o/--orientation`).
 *   - md     : the same aligned table written to a `.md` text file.
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
import PDFDocument from 'pdfkit';
import { ActionSpec, OutputFormat, PageOrientation } from '../parser/types';
import { ExtractConfig } from '../config';
import { ExtractError } from '../errors';
import { ExtractionResult } from '../engine/extract';
import { render, renderAligned, formatFromPath } from './render';

export interface OutputContext {
  config: ExtractConfig;
  /** Sink for stdout (overridable so the test runner can capture output). */
  write: (text: string) => void;
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
      case 'pdf':
        // PDF output is the two-pass aligned table, honoring orientation.
        await writePdf(
          requirePath(action.pdf, 'pdf'),
          renderAligned(results),
          action.orientation,
        );
        break;
      case 'md':
        // The `.md` target writes the same aligned table as plain text.
        writeTextFile(requirePath(action.md, 'md'), renderAligned(results));
        break;
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

/** Render the aligned table text into a simple monospace PDF. */
function writePdf(
  filePath: string,
  text: string,
  orientation?: PageOrientation,
): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      const dir = path.dirname(path.resolve(filePath));
      fs.mkdirSync(dir, { recursive: true });
      const doc = new PDFDocument({ margin: 50, layout: orientation });
      const stream = fs.createWriteStream(filePath);
      stream.on('finish', resolve);
      stream.on('error', reject);
      doc.pipe(stream);
      doc.font('Courier').fontSize(11).text(text, { lineGap: 2 });
      doc.end();
    } catch (err) {
      reject(err);
    }
  });
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
