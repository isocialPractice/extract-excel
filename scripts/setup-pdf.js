#!/usr/bin/env node
/**
 * Postinstall helper: locate Python 3, install openpyxl + reportlab via pip,
 * then cache the executable path so the pdf target works even when python3 is
 * not on PATH.
 *
 * Resolution order at install time:
 *   1. Already on PATH as `python3` or `python`  →  cache and install packages.
 *   2. Already in a standard location  →  cache and install packages.
 *   3. Not found  →  attempt OS-appropriate install, then repeat 1–2.
 *
 * Set SKIP_PDF_SETUP=1 to bypass the install attempt (caching still runs if
 * Python is already present). This script never causes `npm install` to fail;
 * all errors are caught.
 */
'use strict';
const fs           = require('fs');
const path         = require('path');
const { execSync } = require('child_process');

// ── well-known Python 3 paths ─────────────────────────────────────────────────

const COMMON_PATHS = (() => {
  switch (process.platform) {
    case 'win32': {
      const local = process.env.LOCALAPPDATA ?? 'C:\\Users\\Default\\AppData\\Local';
      return [
        path.join(local, 'Programs', 'Python', 'Python313', 'python.exe'),
        path.join(local, 'Programs', 'Python', 'Python312', 'python.exe'),
        path.join(local, 'Programs', 'Python', 'Python311', 'python.exe'),
        path.join(local, 'Programs', 'Python', 'Python310', 'python.exe'),
        'C:\\Python313\\python.exe',
        'C:\\Python312\\python.exe',
        'C:\\Python311\\python.exe',
        'C:\\Python310\\python.exe',
      ];
    }
    case 'darwin':
      return [
        '/opt/homebrew/bin/python3',
        '/usr/local/bin/python3',
      ];
    default:
      return [
        '/usr/bin/python3',
        '/usr/local/bin/python3',
      ];
  }
})();

const CACHE_FILE = path.join(__dirname, '..', '.python-path');

// ── helpers ───────────────────────────────────────────────────────────────────

const log = (msg) => console.log(`[extract-excel] ${msg}`);

/** Return true if `exePath` is Python 3. */
function isPython3(exePath) {
  try {
    const out = execSync(`"${exePath}" --version`, {
      stdio: ['ignore', 'pipe', 'pipe'],
    }).toString();
    return /python 3/i.test(out);
  } catch (e) {
    // Some builds write version to stderr.
    try {
      const out = (e.stderr ?? '').toString();
      return /python 3/i.test(out);
    } catch { return false; }
  }
}

/** Probe PATH for `python3` / `python`; return its full path or null. */
function probeFromPath() {
  const where = process.platform === 'win32' ? 'where' : 'which';
  for (const bin of ['python3', 'python']) {
    try {
      const line = execSync(`${where} ${bin}`, { stdio: ['ignore', 'pipe', 'ignore'] })
        .toString().trim().split(/\r?\n/)[0];
      if (line && fs.existsSync(line) && isPython3(line)) return line;
    } catch { /* not on PATH */ }
  }
  return null;
}

/** Return the first well-known path that resolves to Python 3, or null. */
function probeCommon() {
  for (const p of COMMON_PATHS) {
    if (fs.existsSync(p) && isPython3(p)) return p;
  }
  return null;
}

function findPython() {
  return probeFromPath() ?? probeCommon() ?? null;
}

/** Run a command with visible output; return true on success. */
function tryRun(cmd) {
  try { execSync(cmd, { stdio: 'inherit' }); return true; }
  catch { return false; }
}

/** True if a binary is present at /usr/bin/<name> or on PATH. */
function hasBin(name) {
  if (process.platform !== 'win32' && fs.existsSync(`/usr/bin/${name}`)) return true;
  const where = process.platform === 'win32' ? 'where' : 'which';
  try { execSync(`${where} ${name}`, { stdio: 'ignore' }); return true; }
  catch { return false; }
}

// ── platform installers ───────────────────────────────────────────────────────

function installWindows() {
  log('Trying winget…');
  if (tryRun(
    'winget install --id Python.Python.3.12 --silent ' +
    '--accept-package-agreements --accept-source-agreements',
  )) return true;

  log('winget unavailable or failed; trying Chocolatey…');
  return tryRun('choco install python -y');
}

function installMac() {
  log('Trying Homebrew (brew)…');
  return tryRun('brew install python3');
}

function installLinux() {
  const managers = [
    { bin: 'apt-get', cmd: 'sudo apt-get install -y python3 python3-pip' },
    { bin: 'dnf',     cmd: 'sudo dnf install -y python3 python3-pip' },
    { bin: 'yum',     cmd: 'sudo yum install -y python3 python3-pip' },
    { bin: 'zypper',  cmd: 'sudo zypper install -y python3 python3-pip' },
    { bin: 'pacman',  cmd: 'sudo pacman -S --noconfirm python python-pip' },
  ];
  for (const { bin, cmd } of managers) {
    if (!hasBin(bin)) continue;
    log(`Trying ${bin}…`);
    if (tryRun(cmd)) return true;
  }
  return false;
}

function tryInstallPython() {
  switch (process.platform) {
    case 'win32':  return installWindows();
    case 'darwin': return installMac();
    default:       return installLinux();
  }
}

/** Install openpyxl and reportlab; return true on success. */
function installPackages(python) {
  log('Installing openpyxl and reportlab via pip…');
  try {
    execSync(`"${python}" -m pip install --quiet openpyxl reportlab`, {
      stdio: 'inherit',
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    });
    return true;
  } catch {
    return false;
  }
}

// ── main ──────────────────────────────────────────────────────────────────────

(function main() {
  const skip = process.env.SKIP_PDF_SETUP === '1';
  let found  = findPython();

  if (!found) {
    if (skip) {
      log('SKIP_PDF_SETUP=1 — skipping Python 3 auto-install.');
    } else {
      log('Python 3 not found — attempting automatic installation…');
      const ok = tryInstallPython();
      if (ok) {
        log('Installation finished; locating Python 3…');
        found = findPython();
      } else {
        log('Automatic Python 3 installation did not succeed.');
      }
    }
  }

  if (found) {
    fs.writeFileSync(CACHE_FILE, found, 'utf8');
    log(`Python 3 cached at: ${found}`);
    if (!skip) {
      const ok = installPackages(found);
      if (ok) {
        log('openpyxl + reportlab installed.');
        log('PDF export (--action:pdf) is ready.');
      } else {
        log('pip install failed — run manually:');
        log(`  "${found}" -m pip install openpyxl reportlab`);
      }
    }
  } else {
    try { fs.unlinkSync(CACHE_FILE); } catch { /* no stale cache */ }
    log('');
    log('PDF export will be unavailable until Python 3 is installed.');
    log('Install manually: https://www.python.org/downloads/');
    log('Then run `npm install` (or `npm run setup-pdf`) to cache the path.');
    log('');
    log('After Python 3 is installed, also run:');
    log('  python3 -m pip install openpyxl reportlab');
    log('');
    log('To suppress this message: set SKIP_PDF_SETUP=1 before npm install.');
  }
})();
