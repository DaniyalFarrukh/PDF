'use strict';

/**
 * lib/lo-daemon.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Manages a persistent LibreOffice headless listener (--accept socket mode).
 *
 * LibreOffice is started ONCE at server startup. All conversions connect to
 * this running instance via the UNO Python bridge instead of spawning a new
 * soffice process per conversion.
 *
 * Cold spawn (old):  8-15 s per file
 * UNO daemon (new):  1-3  s per file
 * ─────────────────────────────────────────────────────────────────────────────
 */

const { spawn, execFile } = require('child_process');
const { promisify }       = require('util');
const net  = require('net');
const path = require('path');
const fs   = require('fs');
const os   = require('os');

const execFileAsync = promisify(execFile);

// ── Config ────────────────────────────────────────────────────────────────────
const SOFFICE_PATH  = process.env.SOFFICE_PATH   || 'C:\\Program Files\\LibreOffice\\program\\soffice.exe';
const LO_PYTHON     = process.env.LO_PYTHON_PATH || 'C:\\Program Files\\LibreOffice\\program\\python.exe';
const LO_PORT       = parseInt(process.env.LO_PORT, 10) || 2002;
const DAEMON_DIR    = path.join(os.tmpdir(), 'doctopdf_daemon_profile');
const UNO_SCRIPT    = path.join(__dirname, '..', 'scripts', 'uno_convert.py');
const OUTPUT_DIR    = path.join(__dirname, '..', 'outputs');

// ── Simple serial queue ───────────────────────────────────────────────────────
// UNO Desktop API is single-threaded; serialize all conversions.
class SerialQueue {
  constructor() { this._tail = Promise.resolve(); }
  run(fn) {
    const next = this._tail.then(() => fn()).catch(() => {});
    this._tail = next;
    return next;
  }
  // Public-facing: run fn and surface errors to caller
  add(fn) {
    return new Promise((resolve, reject) => {
      this._tail = this._tail.then(fn).then(resolve, reject).catch(() => {});
    });
  }
}

const queue = new SerialQueue();

// ── Daemon state ──────────────────────────────────────────────────────────────
let _proc         = null;
let _ready        = false;
let _startPromise = null;

// ── Helpers ───────────────────────────────────────────────────────────────────
function isPortOpen(port, timeout = 600) {
  return new Promise(resolve => {
    const s = new net.Socket();
    s.setTimeout(timeout);
    s.on('connect', () => { s.destroy(); resolve(true); });
    s.on('error',   () => { s.destroy(); resolve(false); });
    s.on('timeout', () => { s.destroy(); resolve(false); });
    s.connect(port, '127.0.0.1');
  });
}

async function waitForPort(maxMs = 45_000, interval = 500) {
  const deadline = Date.now() + maxMs;
  while (Date.now() < deadline) {
    if (await isPortOpen(LO_PORT)) return true;
    await new Promise(r => setTimeout(r, interval));
  }
  return false;
}

// ── Start daemon ──────────────────────────────────────────────────────────────
function startDaemon() {
  if (_ready)        return Promise.resolve(true);
  if (_startPromise) return _startPromise;
  _startPromise = _launch();
  return _startPromise;
}

async function _launch() {
  // The LibreOffice --accept socket mode is unreliable on Windows:
  // the process starts but never opens the UNO port, and then blocks
  // all subsequent headless conversions by holding a profile lock.
  // We disable the daemon on Windows and always use the direct-spawn path.
  if (os.platform() === 'win32') {
    console.log('[lo-daemon] Windows detected — UNO daemon disabled; using direct spawn.');
    return false;
  }

  if (!fs.existsSync(SOFFICE_PATH)) {
    console.warn('[lo-daemon] soffice.exe not found — UNO fast-path disabled.');
    return false;
  }
  if (!fs.existsSync(LO_PYTHON)) {
    console.warn('[lo-daemon] LibreOffice python.exe not found — UNO fast-path disabled.');
    return false;
  }

  // Reuse if already listening (e.g. nodemon restart)
  if (await isPortOpen(LO_PORT)) {
    console.log(`[lo-daemon] Found existing listener on port ${LO_PORT}.`);
    _ready = true;
    return true;
  }

  fs.mkdirSync(DAEMON_DIR, { recursive: true });
  const profileUri = 'file:///' + DAEMON_DIR.replace(/\\/g, '/');

  console.log('[lo-daemon] Starting LibreOffice listener…');

  _proc = spawn(SOFFICE_PATH, [
    `--env:UserInstallation=${profileUri}`,
    '--headless',
    '--norestore',
    '--nofirststartwizard',
    '--nologo',
    `--accept=socket,host=127.0.0.1,port=${LO_PORT};urp;StarOffice.ServiceManager`,
  ], { windowsHide: true, stdio: 'ignore', detached: false });

  _proc.on('exit', code => {
    console.warn(`[lo-daemon] Exited (code ${code}). Will restart on next request.`);
    _proc = null; _ready = false; _startPromise = null;
  });
  _proc.on('error', err => {
    console.error('[lo-daemon] Spawn error:', err.message);
    _proc = null; _ready = false; _startPromise = null;
  });

  const ok = await waitForPort(45_000);
  if (ok) {
    console.log(`[lo-daemon] ✓ Ready on port ${LO_PORT}`);
    _ready = true;
  } else {
    console.error('[lo-daemon] Did not become ready within 45 s.');
  }
  return _ready;
}

// ── Convert via UNO ───────────────────────────────────────────────────────────
async function convertViaDaemon(inputPath) {
  return queue.add(async () => {
    if (!_ready) {
      const ok = await startDaemon();
      if (!ok) throw new Error('LibreOffice daemon unavailable.');
    }

    fs.mkdirSync(OUTPUT_DIR, { recursive: true });

    const { stdout, stderr } = await execFileAsync(
      LO_PYTHON, [UNO_SCRIPT, inputPath, OUTPUT_DIR, String(LO_PORT)],
      { timeout: 120_000 }
    );

    const errOut = (stderr || '').trim();
    if (errOut && errOut.includes('UNOERROR')) {
      // If it's a connection error, mark daemon not ready so next call restarts it
      if (errOut.includes('Cannot connect')) { _ready = false; _startPromise = null; }
      throw new Error(errOut);
    }

    const line = (stdout || '').trim();
    if (!line.startsWith('OK:')) throw new Error(`UNO output: ${line || '(empty)'}`);

    const pdfPath = line.slice(3).trim();
    if (!fs.existsSync(pdfPath)) throw new Error('PDF not found after UNO conversion.');
    return pdfPath;
  });
}

// ── Shutdown ──────────────────────────────────────────────────────────────────
function stopDaemon() {
  if (_proc) { _proc.kill(); _proc = null; _ready = false; _startPromise = null; }
}

process.on('exit',    stopDaemon);
process.on('SIGINT',  () => { stopDaemon(); process.exit(0); });
process.on('SIGTERM', () => { stopDaemon(); process.exit(0); });

module.exports = { startDaemon, stopDaemon, convertViaDaemon, isDaemonReady: () => _ready };
