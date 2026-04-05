'use strict';

/**
 * lib/converter.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Two-tier conversion strategy:
 *
 *   Tier 1 — UNO daemon  (~1-3 s): routes through the persistent LibreOffice
 *             listener; no process-spawn overhead.
 *
 *   Tier 2 — Direct spawn (~8-15 s): fallback if the daemon is not yet ready
 *             or a UNO error occurs. Still uses LibreOffice headless; never
 *             falls back to HTML rendering.
 * ─────────────────────────────────────────────────────────────────────────────
 */

const { execFile, execSync } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const fs   = require('fs');
const os   = require('os');

const execFileAsync = promisify(execFile);
const { convertViaDaemon, isDaemonReady } = require('./lo-daemon');

// ── Config ────────────────────────────────────────────────────────────────────
const SOFFICE_PATH = process.env.SOFFICE_PATH || 'C:\\Program Files\\LibreOffice\\program\\soffice.exe';
const OUTPUT_DIR   = path.join(__dirname, '..', 'outputs');
const LO_TIMEOUT   = 180_000;

// ── Helpers ───────────────────────────────────────────────────────────────────
function ensureOutputDir() {
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function libreOfficeAvailable() {
  try { fs.accessSync(SOFFICE_PATH, fs.constants.X_OK); return true; } catch {}
  return fs.existsSync(SOFFICE_PATH);
}

// ── Tier 1: UNO daemon ────────────────────────────────────────────────────────
async function convertWithUno(inputPath, baseName) {
  ensureOutputDir();

  const rawPdf = await convertViaDaemon(inputPath);              // <uuid>.pdf in OUTPUT_DIR
  const want   = path.join(OUTPUT_DIR, `${baseName}.pdf`);

  if (rawPdf !== want && fs.existsSync(rawPdf)) {
    try { fs.renameSync(rawPdf, want); } catch { return rawPdf; }
    return want;
  }
  return rawPdf;
}

// ── Windows: kill stale soffice processes ─────────────────────────────────────
function killStaleLibreOffice() {
  if (os.platform() !== 'win32') return;
  try { execSync('taskkill /F /IM soffice.bin /T', { stdio: 'ignore' }); } catch {}
  try { execSync('taskkill /F /IM soffice.exe /T', { stdio: 'ignore' }); } catch {}
  // Brief wait for the process to release its profile lock
  try { execSync('ping -n 2 127.0.0.1 >nul', { stdio: 'ignore' }); } catch {}
}

// ── Tier 2: Direct spawn ──────────────────────────────────────────────────────
async function convertWithSpawn(inputPath, attempt = 1) {
  // On Windows, kill any stale LibreOffice that would lock the profile
  if (attempt === 1) killStaleLibreOffice();
  ensureOutputDir();

  const stem      = path.basename(inputPath, path.extname(inputPath));
  const outputPdf = path.join(OUTPUT_DIR, `${stem}.pdf`);

  if (fs.existsSync(outputPdf)) { try { fs.unlinkSync(outputPdf); } catch {} }

  // LibreOffice on Windows fails silently when --outdir path contains spaces.
  // Use a short, space-free temp dir, then move the PDF to OUTPUT_DIR.
  const tmpOut = path.join(os.tmpdir(), `lo_out_${process.pid}_${Date.now()}`);
  fs.mkdirSync(tmpOut, { recursive: true });

  // Also copy the input to a space-free temp path if needed
  let actualInput = inputPath;
  if (inputPath.includes(' ')) {
    const tmpIn = path.join(os.tmpdir(), `lo_in_${Date.now()}${path.extname(inputPath)}`);
    fs.copyFileSync(inputPath, tmpIn);
    actualInput = tmpIn;
  }

  // Capture stem before try/finally may delete the temp copy of actualInput
  const tmpStem = path.basename(actualInput, path.extname(actualInput));

  try {
    await new Promise((resolve, reject) => {
      const { spawn } = require('child_process');

      const args = [
        '--headless',
        '--norestore',
        '--nofirststartwizard',
        '--nologo',
        '--convert-to', 'pdf:writer_pdf_Export',
        '--outdir', tmpOut,
        actualInput,
      ];

      console.log('[converter] Spawning LibreOffice:', SOFFICE_PATH, args.join(' '));

      const proc = spawn(SOFFICE_PATH, args, {
        windowsHide: true,
        stdio: 'ignore',
      });

      const timer = setTimeout(() => {
        proc.kill();
        reject(new Error('LibreOffice timed out after 3 minutes.'));
      }, LO_TIMEOUT);

      proc.on('close', code => {
        clearTimeout(timer);
        console.log(`[converter] LibreOffice exited with code ${code}`);
        resolve(code);
      });

      proc.on('error', err => {
        clearTimeout(timer);
        reject(new Error(`Failed to spawn LibreOffice: ${err.message}`));
      });
    });
  } finally {
    // Clean up temp input copy
    if (actualInput !== inputPath) {
      try { fs.unlinkSync(actualInput); } catch {}
    }
  }

  // Wait briefly for file to be fully written
  await new Promise(r => setTimeout(r, 1000));

  // Find any PDF in tmpOut
  const allPdfs = fs.existsSync(tmpOut)
    ? fs.readdirSync(tmpOut).filter(f => f.endsWith('.pdf'))
    : [];

  if (allPdfs.length === 0) {
    try { fs.rmSync(tmpOut, { recursive: true, force: true }); } catch {}
    if (attempt === 1) {
      console.log('[converter] PDF not found on attempt 1, retrying in 3s…');
      await new Promise(r => setTimeout(r, 3000));
      return convertWithSpawn(inputPath, 2);
    }
    throw new Error('LibreOffice ran but produced no PDF. Check the file is a valid Word document.');
  }

  // Move the PDF to OUTPUT_DIR with the correct UUID-based name
  const foundPdf = path.join(tmpOut, allPdfs[0]);
  try {
    fs.copyFileSync(foundPdf, outputPdf);
  } finally {
    try { fs.rmSync(tmpOut, { recursive: true, force: true }); } catch {}
  }

  return outputPdf;
}

// ── Public API ────────────────────────────────────────────────────────────────
async function convertFile(inputPath, ext, baseName) {
  ensureOutputDir();

  if (!libreOfficeAvailable()) {
    throw new Error('LibreOffice is not installed or not found at the configured path.');
  }

  // Fast path — UNO daemon
  if (isDaemonReady()) {
    try {
      console.log('[converter] Tier-1: UNO daemon');
      const pdfPath = await convertWithUno(inputPath, baseName);
      return { pdfPath, engine: 'LibreOffice' };
    } catch (err) {
      console.warn(`[converter] UNO failed (${err.message}) → falling back to spawn`);
    }
  } else {
    console.log('[converter] Daemon not ready yet → using direct spawn');
  }

  // Slow path — direct spawn
  console.log('[converter] Tier-2: direct spawn');
  const spawnPdf = await convertWithSpawn(inputPath);

  // Rename to UUID-based baseName so the route URL matches
  const stem   = path.basename(inputPath, path.extname(inputPath));
  const rawOut = path.join(OUTPUT_DIR, `${stem}.pdf`);
  const want   = path.join(OUTPUT_DIR, `${baseName}.pdf`);
  if (rawOut !== want && fs.existsSync(rawOut)) {
    try { fs.renameSync(rawOut, want); return { pdfPath: want, engine: 'LibreOffice' }; } catch {}
  }

  return { pdfPath: spawnPdf, engine: 'LibreOffice' };
}

module.exports = { convertFile, libreOfficeAvailable };
