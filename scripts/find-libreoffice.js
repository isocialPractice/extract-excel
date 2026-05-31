#!/usr/bin/env node
/**
 * Postinstall helper: auto-install LibreOffice when missing, then cache its
 * absolute path so the pdf target works even when LibreOffice is not on PATH.
 *
 * Resolution order at install time:
 *   1. Already on PATH  →  cache and done.
 *   2. Already in a standard location  →  cache and done.
 *   3. Not found  →  attempt OS-appropriate install, then repeat 1–2.
 *
 * Set SKIP_LIBREOFFICE_INSTALL=1 to bypass the install attempt (caching still
 * runs). This script never causes `npm install` to fail; all errors are caught.
 */
'use strict';
const fs             = require('fs');
const path           = require('path');
const { execSync }   = require('child_process');

// ── well-known paths ──────────────────────────────────────────────────────────

const COMMON_PATHS = (() => {
  switch (process.platform) {
    case 'win32':  return [
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ];
    case 'darwin': return ['/Applications/LibreOffice.app/Contents/MacOS/soffice'];
    default:       return [
      '/usr/bin/libreoffice',
      '/usr/bin/soffice',
      '/usr/lib/libreoffice/program/soffice',
    ];
  }
})();

const CACHE_FILE = path.join(__dirname, '..', '.libreoffice-path');

// ── helpers ───────────────────────────────────────────────────────────────────

const log = (msg) => console.log(`[extract-excel] ${msg}`);

/** Resolve a binary name via PATH; return its full path or null. */
function probeFromPath() {
  const where = process.platform === 'win32' ? 'where' : 'which';
  for (const bin of ['libreoffice', 'soffice']) {
    try {
      const line = execSync(`${where} ${bin}`, { stdio: ['ignore', 'pipe', 'ignore'] })
        .toString().trim().split(/\r?\n/)[0];
      if (line && fs.existsSync(line)) return line;
    } catch { /* not on PATH */ }
  }
  return null;
}

/** Return the first existing well-known path, or null. */
function probeCommon() {
  return COMMON_PATHS.find(p => fs.existsSync(p)) ?? null;
}

function findLibreOffice() {
  return probeFromPath() ?? probeCommon() ?? null;
}

/** Run a command with visible output; return true on success. */
function tryRun(cmd) {
  try { execSync(cmd, { stdio: 'inherit' }); return true; }
  catch { return false; }
}

/** True if a binary exists at /usr/bin/<name> or anywhere on PATH. */
function hasBin(name) {
  if (fs.existsSync(`/usr/bin/${name}`)) return true;
  try { execSync(`which ${name}`, { stdio: 'ignore' }); return true; }
  catch { return false; }
}

// ── platform installers ───────────────────────────────────────────────────────

function installWindows() {
  log('Trying winget…');
  if (tryRun(
    'winget install --id LibreOffice.LibreOffice --silent ' +
    '--accept-package-agreements --accept-source-agreements',
  )) return true;

  log('winget unavailable or failed; trying Chocolatey…');
  return tryRun('choco install libreoffice -y');
}

function installMac() {
  log('Trying Homebrew (brew)…');
  return tryRun('brew install --cask libreoffice');
}

function installLinux() {
  const managers = [
    { bin: 'apt-get', cmd: 'sudo apt-get install -y libreoffice' },
    { bin: 'dnf',     cmd: 'sudo dnf install -y libreoffice' },
    { bin: 'yum',     cmd: 'sudo yum install -y libreoffice' },
    { bin: 'zypper',  cmd: 'sudo zypper install -y libreoffice' },
    { bin: 'pacman',  cmd: 'sudo pacman -S --noconfirm libreoffice-still' },
    { bin: 'snap',    cmd: 'sudo snap install libreoffice' },
  ];
  for (const { bin, cmd } of managers) {
    if (!hasBin(bin)) continue;
    log(`Trying ${bin}…`);
    if (tryRun(cmd)) return true;
  }
  return false;
}

function tryInstall() {
  switch (process.platform) {
    case 'win32':  return installWindows();
    case 'darwin': return installMac();
    default:       return installLinux();
  }
}

// ── main ──────────────────────────────────────────────────────────────────────

(function main() {
  const skip = process.env.SKIP_LIBREOFFICE_INSTALL === '1';
  let found  = findLibreOffice();

  if (!found) {
    if (skip) {
      log('SKIP_LIBREOFFICE_INSTALL=1 — skipping auto-install.');
    } else {
      log('LibreOffice not found — attempting automatic installation…');
      const ok = tryInstall();
      if (ok) {
        log('Installation finished; locating executable…');
        found = findLibreOffice();
      } else {
        log('Automatic installation did not succeed.');
      }
    }
  }

  if (found) {
    fs.writeFileSync(CACHE_FILE, found, 'utf8');
    log(`LibreOffice cached at: ${found}`);
    log('PDF export (--action:pdf) is ready.');
  } else {
    try { fs.unlinkSync(CACHE_FILE); } catch { /* no stale cache */ }
    log('');
    log('PDF export will be unavailable until LibreOffice is installed.');
    log('Install manually: https://www.libreoffice.org/download/libreoffice-still/');
    log('Then run `npm install` to cache the path, or add LibreOffice to your PATH.');
    log('');
    log(`Default locations on ${process.platform}:`);
    COMMON_PATHS.forEach(p => log(`  ${p}`));
    log('');
    log('To suppress this message: set SKIP_LIBREOFFICE_INSTALL=1 before npm install.');
  }
})();
