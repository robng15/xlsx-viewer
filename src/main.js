import * as XLSX from 'xlsx';
import { renderWorksheet } from './renderer.js';

// ─── DOM References ──────────────────────────────────────────────────────────
const dropZone        = document.getElementById('drop-zone');
const fileInput       = document.getElementById('file-input');
const uploadSection   = document.getElementById('upload-section');
const viewerSection   = document.getElementById('viewer-section');
const fileNameDisplay = document.getElementById('file-name-display');
const fileMeta        = document.getElementById('file-meta');
const sheetTabs       = document.getElementById('sheet-tabs');
const tableWrapper    = document.getElementById('table-wrapper');
const loadingIndicator = document.getElementById('loading-indicator');
const clearBtn        = document.getElementById('clear-btn');
const errorToast      = document.getElementById('error-toast');

// ─── App State ────────────────────────────────────────────────────────────────
let workbook     = null;
let activeSheet  = null;

// ─── Entry Points ─────────────────────────────────────────────────────────────

fileInput.addEventListener('change', (e) => {
  const file = e.target.files?.[0];
  if (file) loadFile(file);
  fileInput.value = ''; // reset so same file can be re-selected
});

// Clicking anywhere in drop-zone (but not the label itself) triggers file picker
dropZone.addEventListener('click', (e) => {
  if (e.target !== document.querySelector('label[for="file-input"]')) {
    fileInput.click();
  }
});

dropZone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});

// Drag & drop
dropZone.addEventListener('dragover',  (e) => { e.preventDefault(); dropZone.classList.add('dragging'); });
dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('dragging'));
dropZone.addEventListener('drop',      (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragging');
  const file = e.dataTransfer.files?.[0];
  if (file) loadFile(file);
});

// Also accept drops on the whole page body
document.body.addEventListener('dragover',  (e) => e.preventDefault());
document.body.addEventListener('drop',      (e) => {
  e.preventDefault();
  const file = e.dataTransfer.files?.[0];
  if (file) loadFile(file);
});

clearBtn.addEventListener('click', resetViewer);

// ─── File Loading ─────────────────────────────────────────────────────────────

async function loadFile(file) {
  // Validate extension
  const name = file.name.toLowerCase();
  if (!name.endsWith('.xls') && !name.endsWith('.xlsx')) {
    showError('Unsupported file type. Please upload an XLS or XLSX file.');
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();

    // Read with SheetJS:
    //   cellStyles: true  — parse cell formatting (fonts, fills, borders, alignment)
    //   cellDates:  true  — parse date serial numbers into JS Date objects
    //   cellNF:     true  — keep number format strings
    //   raw:        false — return formatted text in the `w` property
    workbook = XLSX.read(arrayBuffer, {
      type:       'array',
      cellStyles: true,
      cellDates:  true,
      cellNF:     true,
      raw:        false,
    });

    if (!workbook.SheetNames.length) {
      showError('This workbook contains no sheets.');
      return;
    }

    // Show viewer
    uploadSection.classList.add('hidden');
    viewerSection.classList.remove('hidden');

    fileNameDisplay.textContent = file.name;
    fileMeta.textContent =
      `${formatBytes(file.size)}  ·  ${workbook.SheetNames.length} sheet${workbook.SheetNames.length !== 1 ? 's' : ''}`;

    buildTabs();
    switchSheet(workbook.SheetNames[0]);

  } catch (err) {
    console.error('Failed to read file:', err);
    showError('Could not read the spreadsheet. The file may be corrupted or password-protected.');
  }
}

// ─── Tab Rendering ────────────────────────────────────────────────────────────

function buildTabs() {
  sheetTabs.innerHTML = '';
  for (const name of workbook.SheetNames) {
    const tab = document.createElement('button');
    tab.className    = 'sheet-tab';
    tab.textContent  = name;
    tab.role         = 'tab';
    tab.setAttribute('aria-selected', 'false');
    tab.addEventListener('click', () => switchSheet(name));
    sheetTabs.appendChild(tab);
  }
}

function switchSheet(name) {
  if (name === activeSheet) return;
  activeSheet = name;

  // Update tab appearance
  for (const tab of sheetTabs.querySelectorAll('.sheet-tab')) {
    const isActive = tab.textContent === name;
    tab.classList.toggle('active', isActive);
    tab.setAttribute('aria-selected', String(isActive));
  }

  renderSheet(name);
}

// ─── Sheet Rendering ──────────────────────────────────────────────────────────

function renderSheet(name) {
  loadingIndicator.classList.remove('hidden');
  tableWrapper.innerHTML = '';

  // Yield to let the browser paint the spinner before heavy work
  requestAnimationFrame(() => {
    setTimeout(() => {
      try {
        const ws   = workbook.Sheets[name];
        const html = renderWorksheet(ws);
        tableWrapper.innerHTML = html;
      } catch (err) {
        console.error('Render error:', err);
        tableWrapper.innerHTML = '<p class="empty-sheet">Could not render this sheet.</p>';
      } finally {
        loadingIndicator.classList.add('hidden');
      }
    }, 0);
  });
}

// ─── Reset ────────────────────────────────────────────────────────────────────

function resetViewer() {
  workbook    = null;
  activeSheet = null;
  tableWrapper.innerHTML  = '';
  sheetTabs.innerHTML     = '';
  fileNameDisplay.textContent = '';
  fileMeta.textContent    = '';
  viewerSection.classList.add('hidden');
  uploadSection.classList.remove('hidden');
}

// ─── Utilities ────────────────────────────────────────────────────────────────

function formatBytes(bytes) {
  if (bytes < 1024)       return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

let toastTimer;
function showError(msg) {
  errorToast.textContent = msg;
  errorToast.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => errorToast.classList.add('hidden'), 5000);
}
