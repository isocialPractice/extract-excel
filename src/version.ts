/**
 * Resolve the package version for the `--version` / `-v` flag.
 *
 * The value is read from the shipped `package.json` (one level up from the
 * compiled `dist/`), so it always tracks the published version with no build
 * step to keep in sync. A hard-coded fallback keeps the flag working even if
 * the manifest cannot be read.
 */
import * as fs from 'fs';
import * as path from 'path';

/** Fallback used only if `package.json` cannot be read. */
const FALLBACK_VERSION = '2.0.0';

let cached: string | null = null;

/** Return the current package version (memoized). */
export function getVersion(): string {
  if (cached) return cached;
  const candidate = path.resolve(__dirname, '..', 'package.json');
  try {
    const pkg = JSON.parse(fs.readFileSync(candidate, 'utf8')) as {
      version?: string;
    };
    cached = pkg.version ?? FALLBACK_VERSION;
  } catch {
    cached = FALLBACK_VERSION;
  }
  return cached;
}
