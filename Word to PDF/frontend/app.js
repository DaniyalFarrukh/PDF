/**
 * app.js — DocToPDF Frontend
 * ─────────────────────────────────────────────────────────────────────────────
 * State: idle → fileSelected → uploading → converting → success | error
 * ─────────────────────────────────────────────────────────────────────────────
 */
'use strict';

// ── DOM refs ──────────────────────────────────────────────────────────────────
const dropZone          = document.getElementById('dropZone');
const fileInput         = document.getElementById('fileInput');
const browseBtn         = document.getElementById('browseBtn');
const selectedFileEl    = document.getElementById('selectedFile');
const selectedFileName  = document.getElementById('selectedFileName');
const selectedFileSize  = document.getElementById('selectedFileSize');
const removeFileBtn     = document.getElementById('removeFileBtn');
const convertBtn        = document.getElementById('convertBtn');
const progressSection   = document.getElementById('progressSection');
const progressFill      = document.getElementById('progressFill');
const progressLabel     = document.getElementById('progressLabel');
const progressPct       = document.getElementById('progressPct');
const statusMessage     = document.getElementById('statusMessage');
const statusText        = document.getElementById('statusText');
const resultPanel       = document.getElementById('resultPanel');
const resultMeta        = document.getElementById('resultMeta');
const downloadBtn       = document.getElementById('downloadBtn');
const convertAnotherBtn = document.getElementById('convertAnotherBtn');
const previewContainer  = document.getElementById('previewContainer');
const previewTitle      = document.getElementById('previewTitle');
const pdfFrame          = document.getElementById('pdfFrame');
const closePreviewBtn   = document.getElementById('closePreviewBtn');
const engineBadge       = document.getElementById('engineBadge');
const engineLabel       = document.getElementById('engineLabel');
const adBelowPreview    = document.getElementById('adBelowPreview');

// ── Constants ─────────────────────────────────────────────────────────────────
const ALLOWED_EXTS = new Set(['.doc', '.docx', '.txt']);
const MAX_SIZE_MB  = 50;

// ── State ─────────────────────────────────────────────────────────────────────
let selectedFile  = null;
let currentPdfUrl = null;

// ── Utilities ─────────────────────────────────────────────────────────────────
const show = el => el.classList.remove('hidden');
const hide = el => el.classList.add('hidden');

function formatBytes(bytes) {
  if (bytes < 1024)       return `${bytes} B`;
  if (bytes < 1024 ** 2)  return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(2)} MB`;
}

function getExt(filename) {
  const idx = filename.lastIndexOf('.');
  return idx >= 0 ? filename.slice(idx).toLowerCase() : '';
}

function setProgress(pct, label) {
  const rounded = Math.round(pct);
  progressFill.style.width = `${rounded}%`;
  progressPct.textContent  = `${rounded}%`;
  if (label) progressLabel.textContent = label;
  progressSection.setAttribute('aria-valuenow', rounded);
}

// ── Engine Status ─────────────────────────────────────────────────────────────
async function fetchEngineStatus() {
  try {
    const res  = await fetch('/api/status');
    const data = await res.json();

    show(engineBadge);

    if (data.engine === 'LibreOffice') {
      engineBadge.className = 'engine-badge libre';
      engineLabel.textContent = '✓ LibreOffice — perfect fidelity';
    } else {
      engineBadge.className = 'engine-badge error';
      engineLabel.textContent = '✗ LibreOffice not found — conversion unavailable';
    }
  } catch {
    /* server might still be starting */
  }
}

// ── Error / Status Messages ────────────────────────────────────────────────────
function showError(msg) {
  statusText.textContent   = msg;
  statusMessage.className  = 'status-message';
  show(statusMessage);
  setTimeout(() => hide(statusMessage), 10_000);
}

function hideError() { hide(statusMessage); }

// ── File Selection ────────────────────────────────────────────────────────────
function handleFileSelect(file) {
  if (!file) return;
  hideError();

  const ext = getExt(file.name);
  if (!ALLOWED_EXTS.has(ext)) {
    showError(`Unsupported file type "${ext}". Please upload a .doc, .docx, or .txt file.`);
    resetFileInput();
    return;
  }

  if (file.size > MAX_SIZE_MB * 1024 * 1024) {
    showError(`File is too large (${formatBytes(file.size)}). Maximum is ${MAX_SIZE_MB} MB.`);
    resetFileInput();
    return;
  }

  selectedFile = file;
  selectedFileName.textContent = file.name;
  selectedFileSize.textContent = formatBytes(file.size);
  show(selectedFileEl);
  convertBtn.disabled = false;

  // Hide any previous result
  hide(resultPanel);
  hidePreview();
}

function resetFileInput() {
  fileInput.value = '';
  selectedFile    = null;
  hide(selectedFileEl);
  convertBtn.disabled = true;
}

function resetToIdle() {
  resetFileInput();
  hide(progressSection);
  hide(resultPanel);
  convertBtn.classList.remove('loading');
  convertBtn.disabled = true;
  setProgress(0, 'Uploading…');
  hidePreview();
  hideError();
  currentPdfUrl = null;
}

// ── PDF Preview ───────────────────────────────────────────────────────────────
function showPreview(url, filename) {
  previewTitle.textContent = filename;
  pdfFrame.src = url;
  show(previewContainer);
  previewContainer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function hidePreview() {
  hide(previewContainer);
  pdfFrame.src = '';
}

// ── Conversion ────────────────────────────────────────────────────────────────
async function startConversion() {
  if (!selectedFile) return;

  hideError();
  convertBtn.classList.add('loading');
  convertBtn.disabled = true;
  setProgress(0, 'Uploading…');
  show(progressSection);
  hide(resultPanel);
  hidePreview();

  const formData = new FormData();
  formData.append('document', selectedFile);

  return new Promise((resolve) => {
    const xhr = new XMLHttpRequest();

    // Upload progress → 0–70%
    xhr.upload.addEventListener('progress', (e) => {
      if (e.lengthComputable) {
        setProgress((e.loaded / e.total) * 70, 'Uploading document…');
      }
    });

    // Upload done, server converting → 70–95%
    xhr.upload.addEventListener('load', () => {
      setProgress(70, 'Converting with LibreOffice…');
      let cur = 70;
      const iv = setInterval(() => {
        cur = Math.min(cur + 0.4, 95);
        setProgress(cur, 'Converting with LibreOffice…');
        if (cur >= 95) clearInterval(iv);
      }, 250);
      xhr._iv = iv;
    });

    xhr.addEventListener('load', () => {
      if (xhr._iv) clearInterval(xhr._iv);
      convertBtn.classList.remove('loading');

      if (xhr.status >= 200 && xhr.status < 300) {
        let data;
        try { data = JSON.parse(xhr.responseText); } catch {
          showError('Unexpected server response. Please try again.');
          hide(progressSection);
          convertBtn.disabled = false;
          return resolve();
        }

        if (data.success) {
          setProgress(100, 'Complete!');
          currentPdfUrl = data.pdfUrl;

          const dlName = data.originalName
            ? data.originalName.replace(/\.(doc|docx|txt)$/i, '.pdf')
            : data.filename;

          downloadBtn.href     = data.pdfUrl;
          downloadBtn.download = dlName;
          resultMeta.textContent = `Converted via ${data.engine} · ${dlName}`;

          show(resultPanel);
          resultPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
          setTimeout(() => hide(progressSection), 700);

          // Auto-open preview
          showPreview(data.pdfUrl, dlName);

          // Show below-preview ad
          show(adBelowPreview);

        } else {
          showError(data.error || 'Conversion failed. Please try again.');
          hide(progressSection);
          convertBtn.disabled = false;
        }

      } else {
        let msg = 'Conversion failed. Please try again.';
        try {
          const err = JSON.parse(xhr.responseText);
          if (err.error) msg = err.error;
        } catch {}
        showError(msg);
        hide(progressSection);
        convertBtn.disabled = false;
      }

      resolve();
    });

    xhr.addEventListener('error', () => {
      if (xhr._iv) clearInterval(xhr._iv);
      convertBtn.classList.remove('loading');
      hide(progressSection);
      convertBtn.disabled = false;
      showError('Network error. Please check your connection and try again.');
      resolve();
    });

    xhr.addEventListener('timeout', () => {
      if (xhr._iv) clearInterval(xhr._iv);
      convertBtn.classList.remove('loading');
      hide(progressSection);
      convertBtn.disabled = false;
      showError('Request timed out. The file may be large — please try again.');
      resolve();
    });

    xhr.timeout = 300_000; // 5 minutes
    xhr.open('POST', '/api/convert');
    xhr.send(formData);
  });
}

// ── Event Listeners ───────────────────────────────────────────────────────────

fileInput.addEventListener('change', e => {
  if (e.target.files[0]) handleFileSelect(e.target.files[0]);
});

browseBtn.addEventListener('click', e => { e.stopPropagation(); fileInput.click(); });
dropZone.addEventListener('click', () => { if (!selectedFile) fileInput.click(); });
dropZone.addEventListener('keydown', e => {
  if ((e.key === 'Enter' || e.key === ' ') && !selectedFile) {
    e.preventDefault(); fileInput.click();
  }
});

// Drag & Drop
dropZone.addEventListener('dragenter', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', e => {
  if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over');
});
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) handleFileSelect(e.dataTransfer.files[0]);
});

// Prevent browser from opening dropped files elsewhere
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop',     e => e.preventDefault());

removeFileBtn.addEventListener('click', e => { e.stopPropagation(); resetFileInput(); hideError(); });
convertBtn.addEventListener('click', startConversion);
closePreviewBtn.addEventListener('click', hidePreview);
convertAnotherBtn.addEventListener('click', resetToIdle);

// ── Init ──────────────────────────────────────────────────────────────────────
// Hide ad below preview until conversion succeeds
hide(adBelowPreview);

(async () => { await fetchEngineStatus(); })();
