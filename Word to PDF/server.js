'use strict';

/**
 * server.js — Universal Document to PDF Converter
 * ─────────────────────────────────────────────────────────────────────────────
 * Express server that:
 *  - Serves the frontend SPA from /frontend
 *  - Accepts file uploads via POST /api/convert
 *  - Streams generated PDFs via GET /output/:filename
 *  - Enforces file type and size limits
 *  - Schedules automatic cleanup of temp files
 * ─────────────────────────────────────────────────────────────────────────────
 */

require('dotenv').config();

const express = require('express');
const multer  = require('multer');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');
const { v4: uuidv4 } = require('uuid');

const { convertFile, libreOfficeAvailable } = require('./lib/converter');
const { scheduleDelete, purgeOld }          = require('./lib/cleanup');
const { startDaemon, isDaemonReady }         = require('./lib/lo-daemon');

// ── Constants ─────────────────────────────────────────────────────────────────
const PORT         = parseInt(process.env.PORT, 10) || 3000;
const MAX_SIZE_MB  = parseInt(process.env.MAX_FILE_SIZE_MB, 10) || 50;
const UPLOAD_DIR   = path.join(__dirname, 'uploads');
const OUTPUT_DIR   = path.join(__dirname, 'outputs');

const ALLOWED_MIMES = new Set([
  'application/msword',                                                          // .doc
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',    // .docx
  'text/plain',                                                                  // .txt
]);

const ALLOWED_EXTS = new Set(['.doc', '.docx', '.txt']);

// ── Bootstrap ─────────────────────────────────────────────────────────────────
[UPLOAD_DIR, OUTPUT_DIR].forEach(d => fs.mkdirSync(d, { recursive: true }));
purgeOld();

// ── Multer config ─────────────────────────────────────────────────────────────
const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const ext  = path.extname(file.originalname).toLowerCase();
    const stem = uuidv4();
    cb(null, `${stem}${ext}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: MAX_SIZE_MB * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    // Accept based on extension (MIME types can be spoofed / browser-dependent)
    if (ALLOWED_EXTS.has(ext)) return cb(null, true);
    cb(new multer.MulterError('LIMIT_UNEXPECTED_FILE',
      `Unsupported file type "${ext}". Allowed: .doc, .docx, .txt`));
  },
});

// ── App ───────────────────────────────────────────────────────────────────────
const app = express();

app.use(cors());
app.use(express.json());

// Serve frontend static files
app.use(express.static(path.join(__dirname, 'frontend')));

// Serve generated PDFs
app.use('/output', express.static(OUTPUT_DIR));

// ── Routes ────────────────────────────────────────────────────────────────────

/** Health/status check */
app.get('/api/status', (_req, res) => {
  const lo = libreOfficeAvailable();
  res.json({
    status:        'ok',
    engine:        lo ? 'LibreOffice' : 'unavailable',
    daemonReady:   isDaemonReady(),
    maxFileSizeMB: MAX_SIZE_MB,
    allowedTypes:  [...ALLOWED_EXTS],
  });
});

/** Convert endpoint */
app.post('/api/convert', (req, res, next) => {
  upload.single('document')(req, res, async (err) => {
    // ── Multer errors ──
    if (err) {
      if (err instanceof multer.MulterError) {
        const msg = err.code === 'LIMIT_FILE_SIZE'
          ? `File too large. Maximum allowed size is ${MAX_SIZE_MB} MB.`
          : err.message;
        return res.status(400).json({ error: msg });
      }
      return res.status(400).json({ error: err.message });
    }

    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    const inputPath = req.file.path;
    const ext       = path.extname(req.file.originalname).toLowerCase();
    const baseName  = path.basename(req.file.filename, path.extname(req.file.filename));

    console.log(`[server] Convert request: ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB)`);

    try {
      const { pdfPath, engine } = await convertFile(inputPath, ext, baseName);

      const filename = path.basename(pdfPath);
      const pdfUrl   = `/output/${filename}`;

      console.log(`[server] Conversion successful via ${engine}: ${filename}`);

      // Schedule cleanup for both the upload and the output
      scheduleDelete(inputPath);
      scheduleDelete(pdfPath);

      return res.json({
        success: true,
        pdfUrl,
        filename,
        originalName: req.file.originalname,
        engine,
      });
    } catch (convErr) {
      console.error('[server] Conversion error:', convErr);
      // Clean up upload on failure
      try { fs.unlinkSync(inputPath); } catch {}
      return res.status(500).json({
        error: `Conversion failed: ${convErr.message}`,
      });
    }
  });
});

/** Fallback → serve the SPA */
app.get('*', (_req, res) => {
  res.sendFile(path.join(__dirname, 'frontend', 'index.html'));
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  // Kick off the LibreOffice daemon immediately so it is warm when users
  // arrive. startDaemon() is non-blocking — the server accepts requests
  // right away and falls back to direct spawn until the daemon is ready.
  startDaemon().catch(err => console.warn('[server] Daemon start warning:', err.message));
  console.log('');
  console.log('  ╔══════════════════════════════════════════════════╗');
  console.log('  ║     Universal Document → PDF Converter           ║');
  console.log('  ╚══════════════════════════════════════════════════╝');
  console.log('');
  console.log(`  🌐  http://localhost:${PORT}`);
  console.log(`  ⚙️   Engine  : ${libreOfficeAvailable() ? '✅ LibreOffice (headless)' : '⚠️  HTML Pipeline (Puppeteer) — LibreOffice not found'}`);
  console.log(`  📦  Max size : ${MAX_SIZE_MB} MB`);
  console.log(`  ⏱️   Cleanup  : ${parseInt(process.env.OUTPUT_TTL_MS, 10) > 0 ? `${parseInt(process.env.OUTPUT_TTL_MS, 10) / 60000} min TTL` : 'disabled'}`);
  console.log('');
});
