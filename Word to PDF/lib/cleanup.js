'use strict';

/**
 * lib/cleanup.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Schedules automatic deletion of temporary upload and output files after a
 * configurable TTL.  Also exposes a manual purgeOld() for startup cleanup.
 * ─────────────────────────────────────────────────────────────────────────────
 */

const fs = require('fs');
const path = require('path');

const UPLOAD_DIR = path.join(__dirname, '..', 'uploads');
const OUTPUT_DIR = path.join(__dirname, '..', 'outputs');

const TTL_MS = parseInt(process.env.OUTPUT_TTL_MS, 10) || 600_000; // default: 10 min

/**
 * Schedule a single file for deletion after TTL.
 * @param {string} filePath  Absolute path to file
 */
function scheduleDelete(filePath) {
  if (TTL_MS <= 0) return; // disabled
  setTimeout(() => {
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
        console.log(`[cleanup] Deleted: ${path.basename(filePath)}`);
      }
    } catch (err) {
      console.warn(`[cleanup] Could not delete ${filePath}: ${err.message}`);
    }
  }, TTL_MS);
}

/**
 * On server startup, delete any leftover files older than TTL from a previous run.
 */
function purgeOld() {
  const now = Date.now();
  for (const dir of [UPLOAD_DIR, OUTPUT_DIR]) {
    if (!fs.existsSync(dir)) continue;
    for (const file of fs.readdirSync(dir)) {
      const fullPath = path.join(dir, file);
      try {
        const stat = fs.statSync(fullPath);
        if (now - stat.mtimeMs > TTL_MS) {
          fs.unlinkSync(fullPath);
          console.log(`[cleanup] Purged stale file: ${file}`);
        }
      } catch { /* ignore */ }
    }
  }
}

module.exports = { scheduleDelete, purgeOld };
