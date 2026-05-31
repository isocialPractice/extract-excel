/**
 * Runtime configuration loading.
 *
 * The draft doc notes that the column separator ("--") "can be changed in the
 * config.json file". We resolve config from (in order of precedence):
 *   1. An explicit object passed via the API.
 *   2. A `config.json` next to the package install (shipped defaults).
 *   3. Hard-coded defaults below.
 */
import * as fs from 'fs';
import * as path from 'path';

export interface ExtractConfig {
  /** Joins columns within a single extracted range row. Default `" -- "`. */
  columnSeparator: string;
  /** Inserted between distinct extract outputs. Empty string => one blank line. */
  blockSeparator: string;
  /** How the header row for `--title` is located when `row=` is absent. */
  headerDetection: 'first-nonempty';
  /** Shell syntax for `--action:var`. `auto` picks by platform. */
  varExport: 'auto' | 'cmd' | 'powershell' | 'posix';
}

export const DEFAULT_CONFIG: ExtractConfig = {
  columnSeparator: ' -- ',
  blockSeparator: '',
  headerDetection: 'first-nonempty',
  varExport: 'auto',
};

let cached: ExtractConfig | null = null;

/** Load and memoize configuration, merged over defaults. */
export function loadConfig(overrides?: Partial<ExtractConfig>): ExtractConfig {
  if (cached && !overrides) return cached;

  let fileConfig: Partial<ExtractConfig> = {};
  const candidate = path.resolve(__dirname, '..', 'config.json');
  try {
    if (fs.existsSync(candidate)) {
      fileConfig = JSON.parse(fs.readFileSync(candidate, 'utf8'));
    }
  } catch {
    // A malformed config.json should never crash the tool; fall back silently.
    fileConfig = {};
  }

  // Allow the capture helper scripts to force a var-export style regardless of
  // the platform default (e.g. bash on Windows needs `posix`, not `cmd`).
  const envStyle = process.env.EXTRACT_EXCEL_VAR_STYLE;
  const envOverride: Partial<ExtractConfig> =
    envStyle === 'cmd' || envStyle === 'powershell' || envStyle === 'posix' || envStyle === 'auto'
      ? { varExport: envStyle }
      : {};

  const merged: ExtractConfig = {
    ...DEFAULT_CONFIG,
    ...fileConfig,
    ...envOverride,
    ...overrides,
  };
  if (!overrides) cached = merged;
  return merged;
}

/** Reset the memoized config (used by the test runner between scenarios). */
export function resetConfigCache(): void {
  cached = null;
}
