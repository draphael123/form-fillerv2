// ╔═══════════════════════════════════════════════════════════════════╗
// ║  DocuFill – popup.js                                             ║
// ║  Main UI logic for the DocuFill Chrome Extension                 ║
// ╚═══════════════════════════════════════════════════════════════════╝
//
// ┌─────────────────────────────────────────────────────────────┐
// │  TABLE OF CONTENTS                                         │
// │                                                            │
// │  1. CONSTANTS & CONFIGURATION .................. line ~20   │
// │  2. APPLICATION STATE .......................... line ~40   │
// │  3. UTILITIES .................................. line ~65   │
// │  4. INITIALIZATION ............................. line ~110  │
// │  5. STORAGE .................................... line ~145  │
// │  6. THEME & SETTINGS ........................... line ~185  │
// │  7. DATA TAB (CSV/XLSX upload & parsing) ....... line ~280  │
// │  8. TEMPLATE MANAGEMENT ........................ line ~470  │
// │  9. MAPPING TAB (field capture & editing) ...... line ~600  │
// │  10. FILL TAB (autofill & history) ............. line ~740  │
// │  11. UI REFRESH & RENDERING .................... line ~950  │
// │  12. DOCUSIGN DETECTION ........................ line ~1120 │
// │  13. MESSAGING ................................. line ~1160 │
// └─────────────────────────────────────────────────────────────┘


// ═════════════════════════════════════════════════════════════════════
// 1. CONSTANTS & CONFIGURATION
// ═════════════════════════════════════════════════════════════════════

const DEFAULT_SETTINGS = {
  theme: 'dark',
  trimWhitespace: true,
  dateFormat: '',
  encoding: 'utf-8',
  autoLoadTemplate: false,
  urlRules: '',
  warnEmpty: true,
  showPreview: true,
  batchFillMode: false,
  switchTabAfterFill: false,
};

const RECORD_PAGE_SIZE = 50;
const FILL_HISTORY_MAX = 10;

// Default source: provider compliance dashboard workflow (see CURSOR_PROMPT / first-use workflow)
const SAMPLE_CSV =
  'Provider Name,NPI,License Number,Email,Agreement Status\nJane Doe,1234567890,UT-12345,jane@example.com,Signed\nJohn Smith,0987654321,UT-67890,john@example.com,Pending';


// ═════════════════════════════════════════════════════════════════════
// 2. APPLICATION STATE
// ═════════════════════════════════════════════════════════════════════

let state = {
  // Data
  csvData: null,
  pendingXlsxData: null,
  pendingXlsxSheets: null,
  pendingXlsxFileName: null,

  // Templates & mappings
  templates: {},
  activeTemplate: null,
  pendingMappings: [],
  pendingField: null,
  captureMode: false,

  // Fill state
  filteredIndices: null,
  filledRecordIndices: [],
  selectedRecord: null,
  recordPage: 0,
  lastFillResult: null,
  lastRestoreData: null,
  fillHistory: [],

  // UI & environment
  isOnDocuSign: false,
  contentScriptReady: false,
  docuSignTabCount: 0,
  lastVerifyResult: null,
  onboardingDismissed: false,
  settings: { ...DEFAULT_SETTINGS },
};


// ═════════════════════════════════════════════════════════════════════
// 3. UTILITIES
// ═════════════════════════════════════════════════════════════════════

const $ = (id) => document.getElementById(id);

function escapeHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

let toastTimeout;
function showToast(msg, type = '') {
  const toast = $('toast');
  toast.textContent = msg;
  toast.className = 'toast show' + (type ? ' ' + type : '');
  clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => toast.classList.remove('show'), 2500);
}

function normalizeForMatch(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[^a-z0-9\s]/g, '');
}

function scoreColumnMatch(fieldLabel, column) {
  if (!column) return 0;
  const a = normalizeForMatch(fieldLabel);
  const b = normalizeForMatch(column);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (fieldLabel.toLowerCase() === column.toLowerCase()) return 0.98;
  if (a.replace(/\s/g, '') === b.replace(/\s/g, '')) return 0.95;
  if (b.includes(a) || a.includes(b)) return 0.7;

  const aWords = a.split(/\s+/).filter(Boolean);
  const bWords = b.split(/\s+/).filter(Boolean);
  const overlap = aWords.filter(
    (w) => bWords.some((bw) => bw.includes(w) || w.includes(bw))
  ).length;
  if (overlap > 0) return 0.4 + (overlap / Math.max(aWords.length, bWords.length)) * 0.4;

  const synonymPairs = [
    ['name', 'provider name'], ['name', 'applicant name'], ['name', 'first name'], ['name', 'full name'],
    ['license', 'license number'], ['license', 'dopl license'], ['number', 'license number'], ['number', 'npi'],
    ['email', 'email address'], ['phone', 'phone number'], ['address', 'practice establishment'], ['status', 'agreement status']
  ];
  for (let i = 0; i < synonymPairs.length; i++) {
    const [x, y] = synonymPairs[i];
    if ((a.includes(x) && b.includes(y)) || (a.includes(y) && b.includes(x))) return 0.65;
  }
  return 0;
}

function confidenceBadge(score) {
  if (score == null) return '';
  const pct = Math.round(score * 100);
  let color;
  if (score >= 0.95) color = 'var(--success)';
  else if (score >= 0.7) color = 'var(--accent)';
  else color = 'var(--warning)';
  return ` <span style="font-size:9px;padding:1px 5px;border-radius:3px;background:${color}22;color:${color};font-weight:600;margin-left:4px;">${pct}%</span>`;
}

function getRecordValue(record, column, mapping) {
  let val = record[column];
  const isEmpty = val === undefined || val === null || String(val).trim() === '';
  if (isEmpty && mapping && mapping.default) return String(mapping.default);
  if (isEmpty) return '';

  val = String(val);
  if (state.settings.trimWhitespace) val = val.trim();

  if (mapping && mapping.transform) {
    if (mapping.transform === 'uppercase') val = val.toUpperCase();
    else if (mapping.transform === 'lowercase') val = val.toLowerCase();
    else if (mapping.transform === 'trim') val = val.trim();
  }

  const dateLike = (s) => /date|dob|birth|year|effective|expir|signed/i.test(String(s || ''));
  if (state.settings.dateFormat && (dateLike(column) || (mapping && dateLike(mapping.fieldLabel)))) {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      const pad = (n) => String(n).padStart(2, '0');
      const y = d.getFullYear();
      const m = pad(d.getMonth() + 1);
      const day = pad(d.getDate());
      if (state.settings.dateFormat === 'MM/DD/YYYY') val = `${m}/${day}/${y}`;
      else if (state.settings.dateFormat === 'DD/MM/YYYY') val = `${day}/${m}/${y}`;
      else if (state.settings.dateFormat === 'YYYY-MM-DD') val = `${y}-${m}-${day}`;
    }
  }
  return val;
}


// ═════════════════════════════════════════════════════════════════════
// 4. INITIALIZATION
// ═════════════════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', async () => {
  try {
    const manifest = chrome.runtime.getManifest();
    const v = $('headerVersion');
    if (v && manifest.version) v.textContent = 'v' + manifest.version;
  } catch (e) {}

  await loadFromStorage();
  if (!state.csvData) loadSampleCsv(true);

  applyTheme(state.settings.theme);
  setupTabs();
  setupSettingsTab();
  setupHelpTab();
  setupOnboarding();
  setupDataTab();
  setupMappingTab();
  setupFillTab();
  checkDocuSignTab();
  listenForContentMessages();
  tryApplyUrlTemplateRules();
});

function setupTabs() {
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach((t) => {
        t.classList.remove('active');
        t.setAttribute('aria-selected', 'false');
      });
      document.querySelectorAll('.panel').forEach((p) => p.classList.remove('active'));
      tab.classList.add('active');
      tab.setAttribute('aria-selected', 'true');
      const panel = $('tab-' + tab.dataset.tab);
      if (panel) panel.classList.add('active');
    });
  });
}

function setupOnboarding() {
  if (!state.onboardingDismissed) {
    const overlay = $('onboardingOverlay');
    if (overlay) overlay.style.display = 'flex';
  }
  $('onboardingDismiss').addEventListener('click', () => {
    state.onboardingDismissed = true;
    saveToStorage();
    const overlay = $('onboardingOverlay');
    if (overlay) overlay.style.display = 'none';
    showToast('You can reopen Help anytime');
  });
}

function setupHelpTab() {}


// ═════════════════════════════════════════════════════════════════════
// 5. STORAGE
// ═════════════════════════════════════════════════════════════════════

async function loadFromStorage() {
  return new Promise((resolve) => {
    const keys = [
      'csvData', 'templates', 'lastTemplate', 'lastRecord',
      'settings', 'onboardingDismissed', 'fillHistory', 'filledRecordIndices',
    ];
    chrome.storage.local.get(keys, async (result) => {
      // Migrate csvData from chrome.storage.local to IndexedDB (one-time)
      if (result.csvData) {
        state.csvData = result.csvData;
        try {
          await saveCsvData(result.csvData);
          chrome.storage.local.remove('csvData');
        } catch (e) {
          console.error('IndexedDB migration failed:', e);
        }
      } else {
        try {
          const idbData = await loadCsvData();
          if (idbData) state.csvData = idbData;
        } catch (e) {
          console.error('IndexedDB load failed:', e);
        }
      }

      if (result.templates) state.templates = result.templates;
      if (result.settings && typeof result.settings === 'object') {
        state.settings = { ...DEFAULT_SETTINGS, ...result.settings };
      }
      if (result.onboardingDismissed) state.onboardingDismissed = true;
      if (Array.isArray(result.fillHistory)) {
        state.fillHistory = result.fillHistory.slice(0, FILL_HISTORY_MAX);
      }
      if (Array.isArray(result.filledRecordIndices)) {
        state.filledRecordIndices = result.filledRecordIndices;
      }

      refreshAll();

      if (result.lastTemplate && state.templates[result.lastTemplate]) {
        $('templateSelect').value = result.lastTemplate;
        state.selectedRecord = null;
      }
      if (
        result.lastRecord != null &&
        state.csvData &&
        result.lastRecord >= 0 &&
        result.lastRecord < state.csvData.rows.length
      ) {
        state.selectedRecord = result.lastRecord;
        state.recordPage = Math.floor(result.lastRecord / RECORD_PAGE_SIZE);
        $('recordSelect').value = String(result.lastRecord);
      }

      updateFillButton();
      refreshSettingsUI();

      // Auto-load template by URL pattern
      if (state.settings.autoLoadTemplate && state.settings.urlRules && Object.keys(state.templates).length > 0) {
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
          const url = (tabs[0] && tabs[0].url) || '';
          const lines = state.settings.urlRules.split('\n').map((l) => l.trim()).filter(Boolean);
          for (const line of lines) {
            const parts = line.split(/\s+/);
            const pattern = parts[0];
            const templateName = parts.slice(1).join(' ').trim();
            if (pattern && templateName && state.templates[templateName] && url.indexOf(pattern) >= 0) {
              const sel = $('templateSelect');
              if (sel) sel.value = templateName;
              saveToStorage();
              break;
            }
          }
          resolve();
        });
      } else {
        resolve();
      }
    });
  });
}

function saveToStorage() {
  const payload = {
    templates: state.templates,
    settings: state.settings,
    onboardingDismissed: state.onboardingDismissed,
    fillHistory: state.fillHistory.slice(0, FILL_HISTORY_MAX),
    filledRecordIndices: state.filledRecordIndices,
  };

  const lastTemplate = $('templateSelect')?.value || '';
  if (lastTemplate) payload.lastTemplate = lastTemplate;

  if (state.csvData && state.selectedRecord != null) {
    payload.lastRecord = state.selectedRecord;
  } else {
    payload.lastRecord = null;
  }

  chrome.storage.local.set(payload);

  // Save csvData to IndexedDB (fire-and-forget)
  if (state.csvData) {
    saveCsvData(state.csvData).catch((e) => console.error('IndexedDB save failed:', e));
  }
}


// ═════════════════════════════════════════════════════════════════════
// 6. THEME & SETTINGS
// ═════════════════════════════════════════════════════════════════════

function applyTheme(theme) {
  if (theme === 'system') {
    const m = window.matchMedia('(prefers-color-scheme: dark)');
    document.body.classList.toggle('theme-light', !m.matches);
  } else {
    document.body.classList.toggle('theme-light', theme === 'light');
  }
}

function refreshSettingsUI() {
  const s = state.settings;

  const setVal = (id, val) => { const el = $(id); if (el) el.value = val; };
  const setChecked = (id, val) => { const el = $(id); if (el) el.checked = val; };

  setVal('settingTheme', s.theme || 'dark');
  setChecked('settingTrimWhitespace', s.trimWhitespace !== false);
  setVal('settingDateFormat', s.dateFormat || '');
  setVal('settingEncoding', s.encoding || 'utf-8');
  setChecked('settingAutoLoadTemplate', !!s.autoLoadTemplate);
  setVal('settingUrlRules', s.urlRules || '');
  setChecked('settingWarnEmpty', s.warnEmpty !== false);
  setChecked('settingShowPreview', s.showPreview !== false);
  setChecked('settingBatchFillMode', !!s.batchFillMode);
  setChecked('settingSwitchTabAfterFill', !!s.switchTabAfterFill);

  const wrap = $('urlRulesWrap');
  const autoEl = $('settingAutoLoadTemplate');
  if (wrap) wrap.style.display = autoEl && autoEl.checked ? 'block' : 'none';
}

function setupSettingsTab() {
  const settingIds = [
    'settingTheme', 'settingTrimWhitespace', 'settingDateFormat',
    'settingEncoding', 'settingAutoLoadTemplate', 'settingUrlRules',
    'settingWarnEmpty', 'settingShowPreview', 'settingBatchFillMode', 'settingSwitchTabAfterFill',
  ];

  function saveSettings() {
    state.settings = {
      theme: $('settingTheme')?.value || 'dark',
      trimWhitespace: $('settingTrimWhitespace')?.checked !== false,
      dateFormat: $('settingDateFormat')?.value || '',
      encoding: $('settingEncoding')?.value || 'utf-8',
      autoLoadTemplate: $('settingAutoLoadTemplate')?.checked || false,
      urlRules: $('settingUrlRules')?.value || '',
      warnEmpty: $('settingWarnEmpty')?.checked !== false,
      showPreview: $('settingShowPreview')?.checked !== false,
      batchFillMode: $('settingBatchFillMode')?.checked || false,
      switchTabAfterFill: $('settingSwitchTabAfterFill')?.checked || false,
    };
    applyTheme(state.settings.theme);
    $('urlRulesWrap').style.display = state.settings.autoLoadTemplate ? 'block' : 'none';
    saveToStorage();
    showToast('Settings saved');
  }

  settingIds.forEach((id) => {
    $(id).addEventListener('change', saveSettings);
  });
}

function tryApplyUrlTemplateRules() {
  if (!state.settings.autoLoadTemplate || !state.settings.urlRules) return;

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const url = tabs[0]?.url || '';
    const lines = state.settings.urlRules
      .split('\n')
      .map((l) => l.trim())
      .filter(Boolean);

    for (const line of lines) {
      const m = line.match(/^(\S+)\s+(.+)$/);
      if (!m) continue;

      let pattern = m[1];
      const templateName = m[2].trim();
      if (!pattern || !templateName) continue;
      if (pattern.startsWith('/') && pattern.endsWith('/')) {
        pattern = pattern.slice(1, -1);
      }

      try {
        const re = new RegExp(pattern);
        if (re.test(url) && state.templates[templateName]) {
          $('templateSelect').value = templateName;
          saveToStorage();
          updateFillButton();
          break;
        }
      } catch (e) {}
    }
  });
}


// ═════════════════════════════════════════════════════════════════════
// 7. DATA TAB (CSV/XLSX upload & parsing)
// ═════════════════════════════════════════════════════════════════════

function setupDataTab() {
  const uploadZone = $('uploadZone');
  const fileInput = $('fileInput');

  uploadZone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => handleFileUpload(e.target.files[0]));

  uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.classList.add('drag-over');
  });
  uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
  uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFileUpload(e.dataTransfer.files[0]);
  });

  $('clearFileBtn').addEventListener('click', () => {
    state.csvData = null;
    state.selectedRecord = null;
    state.filteredIndices = null;
    state.filledRecordIndices = [];
    saveToStorage();
    clearCsvData().catch((e) => console.error('IndexedDB clear failed:', e));
    refreshAll();
    showToast('File removed');
  });

  $('loadSheetBtn').addEventListener('click', loadSelectedSheet);
  $('cancelSheetBtn').addEventListener('click', () => {
    state.pendingXlsxData = null;
    state.pendingXlsxSheets = null;
    state.pendingXlsxFileName = null;
    refreshAll();
  });

  $('exportTemplatesBtn').addEventListener('click', exportTemplates);
  $('importTemplatesBtn').addEventListener('click', () => $('importTemplatesInput').click());
  $('importTemplatesInput').addEventListener('change', (e) => {
    if (e.target.files[0]) importTemplates(e.target.files[0]);
    e.target.value = '';
  });
  $('importMappingsCsvBtn').addEventListener('click', () => $('importMappingsCsvInput').click());
  $('importMappingsCsvInput').addEventListener('change', (e) => {
    if (e.target.files[0]) importMappingsCsv(e.target.files[0]);
    e.target.value = '';
  });
  $('loadSampleCsvBtn').addEventListener('click', loadSampleCsv);
  $('pasteCsvBtn').addEventListener('click', pasteCsvFromClipboard);
}

async function pasteCsvFromClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    if (!text || !text.trim()) {
      showToast('Clipboard is empty', 'error');
      return;
    }
    showToast('Pasting…');
    const parsed = parseCSV(text);
    state.csvData = { name: 'Pasted', columns: parsed.columns, rows: parsed.rows };
    state.filteredIndices = null;
    state.filledRecordIndices = [];
    saveToStorage();
    refreshAll();
    showToast('Pasted ' + parsed.rows.length + ' row(s)', 'success');
  } catch (e) {
    showToast('Paste failed — try pasting into a text field first', 'error');
  }
}

// ── File handling ───────────────────────────────────────

async function handleFileUpload(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();

  try {
    let rows, columns, duplicateColumns = null;

    if (ext === 'csv') {
      showToast('Parsing file…');
      await new Promise((r) => setTimeout(r, 50));
      const parsed = parseCSV(await file.text());
      rows = parsed.rows;
      columns = parsed.columns;
      duplicateColumns = parsed.duplicateColumns || null;
    } else if (ext === 'xlsx' || ext === 'xls') {
      const parsed = await parseXLSX(file);
      if (parsed.needSheetChoice && parsed.sheetNames && parsed.array) {
        state.pendingXlsxData = parsed.array;
        state.pendingXlsxSheets = parsed.sheetNames;
        state.pendingXlsxFileName = file.name;
        refreshAll();
        showToast('Select a sheet below');
        return;
      }
      rows = parsed.rows;
      columns = parsed.columns;
      duplicateColumns = parsed.duplicateColumns || null;
    } else {
      showToast('Unsupported file type', 'error');
      return;
    }

    if (rows.length === 0) {
      showToast('File has no data rows (only headers). Columns loaded.', 'error');
    }
    if (duplicateColumns && duplicateColumns.length > 0) {
      showToast('Duplicate column names: ' + duplicateColumns.join(', ') + '. Consider renaming.', 'error');
    }

    state.csvData = { name: file.name, columns, rows };
    state.filteredIndices = null;
    state.filledRecordIndices = [];
    saveToStorage();
    refreshAll();
    showToast(
      rows.length ? `Loaded ${rows.length} records` : 'Columns loaded (0 rows)',
      rows.length ? 'success' : ''
    );
  } catch (err) {
    console.error(err);
    showToast('Failed to parse file', 'error');
  }
}

async function loadSelectedSheet() {
  const sel = $('sheetSelect');
  if (!state.pendingXlsxData || !state.pendingXlsxSheets || sel.value === '') return;

  try {
    const result = await parseXLSXWithSheet(state.pendingXlsxData, parseInt(sel.value, 10));
    const { rows, columns, duplicateColumns } = result;

    if (duplicateColumns && duplicateColumns.length > 0) {
      showToast('Duplicate column names: ' + duplicateColumns.join(', ') + '. Consider renaming.', 'error');
    }

    state.csvData = { name: state.pendingXlsxFileName || 'Sheet', columns, rows };
    state.filledRecordIndices = [];
    state.pendingXlsxData = null;
    state.pendingXlsxSheets = null;
    state.pendingXlsxFileName = null;
    saveToStorage();
    refreshAll();
    showToast(
      rows.length ? `Loaded ${rows.length} records` : 'Columns loaded (0 rows)',
      rows.length ? 'success' : ''
    );
  } catch (err) {
    showToast(err.message || 'Failed to load sheet', 'error');
  }
}

function loadSampleCsv(silent) {
  try {
    const parsed = parseCSV(SAMPLE_CSV);
    state.csvData = { name: 'docufill_providers.csv', columns: parsed.columns, rows: parsed.rows };
    state.filledRecordIndices = [];
    saveToStorage();
    refreshAll();
    if (!silent) showToast('Loaded sample CSV (' + parsed.rows.length + ' rows)', 'success');
  } catch (e) {
    if (!silent) showToast('Failed to load sample', 'error');
  }
}

// ── CSV parsing ─────────────────────────────────────────

function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1); // strip BOM
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 1) throw new Error('CSV is empty');

  const columns = splitCSVLine(lines[0]);
  const rows = lines.length < 2
    ? []
    : lines.slice(1).map((line) => {
        const vals = splitCSVLine(line);
        const obj = {};
        columns.forEach((col, i) => { obj[col] = vals[i] ?? ''; });
        return obj;
      });

  const duplicates = columns.filter((c, i) => columns.indexOf(c) !== i);
  const uniqueDupes = [...new Set(duplicates)];
  return { columns, rows, duplicateColumns: uniqueDupes.length ? uniqueDupes : null };
}

function splitCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') inQuotes = !inQuotes;
    else if (ch === ',' && !inQuotes) { result.push(current.trim()); current = ''; }
    else current += ch;
  }
  result.push(current.trim());
  return result;
}

// ── XLSX parsing (via background worker) ────────────────

function sendParseXLSX(dataArray, sheetIndex) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(
      { type: 'PARSE_XLSX', data: dataArray, sheetIndex },
      (response) => {
        if (response && response.success) {
          if (response.needSheetChoice) {
            resolve({ needSheetChoice: true, sheetNames: response.sheetNames });
          } else {
            resolve({ result: response.result });
          }
        } else {
          reject(new Error(response?.error || 'Parse failed'));
        }
      }
    );
  });
}

async function parseXLSX(file) {
  const buf = await file.arrayBuffer();
  const arr = Array.from(new Uint8Array(buf));
  const out = await sendParseXLSX(arr);
  if (out.needSheetChoice) {
    return { needSheetChoice: true, sheetNames: out.sheetNames, array: arr };
  }
  return out.result;
}

async function parseXLSXWithSheet(dataArray, sheetIndex) {
  const out = await sendParseXLSX(dataArray, sheetIndex);
  if (out.needSheetChoice) throw new Error('Unexpected needSheetChoice');
  return out.result;
}


// ═════════════════════════════════════════════════════════════════════
// 8. TEMPLATE MANAGEMENT
// ═════════════════════════════════════════════════════════════════════

function saveTemplate() {
  const name = $('templateNameInput').value.trim();
  if (!name) { showToast('Enter a template name', 'error'); return; }
  if (state.pendingMappings.length === 0) { showToast('Add at least one mapping', 'error'); return; }

  state.templates[name] = [...state.pendingMappings];
  state.activeTemplate = name;
  saveToStorage();
  refreshFillTemplates();
  refreshSavedTemplates();
  showToast(`Template "${name}" saved`, 'success');
}

function newTemplate() {
  state.pendingMappings = [];
  state.activeTemplate = null;
  $('templateNameInput').value = '';
  renderMappingsList();
}

function loadTemplateIntoEditor() {
  const name = $('templateNameInput').value.trim();
  if (!name || !state.templates[name]) { showToast('Template not found', 'error'); return; }
  state.pendingMappings = [...state.templates[name]];
  state.activeTemplate = name;
  renderMappingsList();
  showToast(`Loaded "${name}"`);
}

function exportTemplates() {
  const json = JSON.stringify(state.templates, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'docufill-templates.json';
  a.click();
  URL.revokeObjectURL(url);
  showToast('Templates exported');
}

async function importTemplates(file) {
  try {
    const text = await file.text();
    const imported = JSON.parse(text);
    if (typeof imported !== 'object' || imported === null) throw new Error('Invalid format');

    let merged = 0;
    Object.keys(imported).forEach((k) => {
      if (Array.isArray(imported[k])) {
        state.templates[k] = imported[k];
        merged++;
      }
    });

    saveToStorage();
    refreshAll();
    showToast(
      merged ? `Imported ${merged} template(s)` : 'No valid templates in file',
      merged ? 'success' : 'error'
    );
  } catch (err) {
    showToast('Import failed: ' + (err.message || 'invalid file'), 'error');
  }
}

async function importMappingsCsv(file) {
  try {
    const text = await file.text();
    const lines = text.trim().split(/\r?\n/).filter(Boolean);
    if (lines.length < 2) { showToast('CSV needs header + at least 1 row', 'error'); return; }

    const cols = splitCSVLine(lines[0]);
    const fieldKeyIdx = cols.findIndex((c) => /fieldkey|field_key/i.test(c));
    const fieldLabelIdx = cols.findIndex((c) => /fieldlabel|field_label/i.test(c));
    const columnIdx = cols.findIndex((c) => /column/i.test(c));
    if (columnIdx < 0) { showToast('CSV must have "column" header', 'error'); return; }

    const keyIdx = fieldKeyIdx >= 0 ? fieldKeyIdx : (cols[0] ? 0 : -1);
    const labelIdx = fieldLabelIdx >= 0 ? fieldLabelIdx : (keyIdx >= 0 ? keyIdx : 0);

    let added = 0;
    for (let i = 1; i < lines.length; i++) {
      const vals = splitCSVLine(lines[i]);
      const fieldKey = vals[keyIdx]?.trim();
      const column = vals[columnIdx]?.trim();
      if (!fieldKey || !column) continue;
      const fieldLabel = vals[labelIdx]?.trim() || fieldKey;
      state.pendingMappings.push({ fieldKey, fieldLabel, column });
      added++;
    }

    renderMappingsList();
    saveToStorage();
    showToast('Imported ' + added + ' mapping(s)', 'success');
  } catch (e) {
    showToast('Import failed: ' + (e.message || 'invalid file'), 'error');
  }
}

function refreshFillTemplates() {
  const sel = $('templateSelect');
  const names = Object.keys(state.templates);
  sel.innerHTML =
    '<option value="">— Select a template —</option>' +
    names
      .map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)} (${state.templates[n].length} fields)</option>`)
      .join('');
}

function refreshSavedTemplates() {
  const list = $('savedTemplatesList');
  const names = Object.keys(state.templates);
  $('templateBadge').textContent = `${names.length} saved`;

  if (names.length === 0) {
    list.innerHTML = '<div class="empty-state" style="padding:16px;">' +
      '<div class="empty-text">No templates saved yet</div></div>';
    return;
  }

  list.innerHTML = names
    .map((name) => `
      <div class="template-item">
        <div>
          <div class="template-name">${escapeHtml(name)}</div>
          <div class="template-meta">${state.templates[name].length} field mappings</div>
        </div>
        <div class="template-actions">
          <button class="btn btn-secondary btn-sm" data-load="${escapeHtml(name)}" title="Edit in Mapping tab">Edit</button>
          <button class="btn btn-secondary btn-sm" data-clone="${escapeHtml(name)}" title="Duplicate template">Clone</button>
          <button class="btn btn-secondary btn-sm" data-rename="${escapeHtml(name)}" title="Rename template">Rename</button>
          <button class="btn btn-danger btn-sm" data-delete="${escapeHtml(name)}" title="Delete template">✕</button>
        </div>
      </div>
    `)
    .join('');

  // Edit
  list.querySelectorAll('[data-load]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.pendingMappings = [...state.templates[btn.dataset.load]];
      $('templateNameInput').value = btn.dataset.load;
      renderMappingsList();
      document.querySelector('.tab[data-tab="mapping"]').click();
    });
  });

  // Clone
  list.querySelectorAll('[data-clone]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const base = btn.dataset.clone;
      let newName = 'Copy of ' + base;
      let c = 1;
      while (state.templates[newName]) {
        newName = 'Copy of ' + base + ' (' + (++c) + ')';
      }
      state.templates[newName] = [...state.templates[base]];
      saveToStorage();
      refreshSavedTemplates();
      refreshFillTemplates();
      showToast('Cloned as "' + newName + '"', 'success');
    });
  });

  // Rename
  list.querySelectorAll('[data-rename]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const oldName = btn.dataset.rename;
      const newName = window.prompt('New template name:', oldName);
      if (!newName || newName.trim() === '') return;
      const trimmed = newName.trim();
      if (trimmed === oldName) return;
      if (state.templates[trimmed]) {
        showToast('A template with that name already exists', 'error');
        return;
      }
      state.templates[trimmed] = state.templates[oldName];
      delete state.templates[oldName];
      if ($('templateSelect').value === oldName) $('templateSelect').value = trimmed;
      if (state.activeTemplate === oldName) state.activeTemplate = trimmed;
      saveToStorage();
      refreshSavedTemplates();
      refreshFillTemplates();
      showToast('Renamed to "' + trimmed + '"', 'success');
    });
  });

  // Delete
  list.querySelectorAll('[data-delete]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const name = btn.dataset.delete;
      if (!window.confirm('Delete template "' + name + '"?')) return;
      delete state.templates[name];
      saveToStorage();
      refreshSavedTemplates();
      refreshFillTemplates();
      showToast(`Deleted "${name}"`);
    });
  });
}


// ═════════════════════════════════════════════════════════════════════
// 9. MAPPING TAB (field capture & editing)
// ═════════════════════════════════════════════════════════════════════

function setupMappingTab() {
  $('scanFieldsBtn').addEventListener('click', scanPageForFields);
  $('captureModeBtn').addEventListener('click', toggleCaptureMode);
  $('confirmMappingBtn').addEventListener('click', confirmMapping);
  $('cancelCaptureBtn').addEventListener('click', cancelCapture);
  $('saveTemplateBtn').addEventListener('click', saveTemplate);
  $('verifyMappingsBtn').addEventListener('click', verifyMappingsOnPage);
  $('loadPresetProviderBtn').addEventListener('click', loadPresetProviderCompliance);
  $('newTemplateBtn').addEventListener('click', newTemplate);
  $('loadTemplateBtn').addEventListener('click', loadTemplateIntoEditor);
}

function loadPresetProviderCompliance() {
  const cols = ['Provider Name', 'NPI', 'License Number', 'Email', 'Agreement Status'];
  const existing = state.pendingMappings.length;
  cols.forEach((name) => {
    if (state.pendingMappings.some((m) => (m.fieldKey || m.fieldLabel) === name)) return;
    state.pendingMappings.push({ fieldKey: name, fieldLabel: name, column: (state.csvData && state.csvData.columns.includes(name)) ? name : (state.csvData && state.csvData.columns[0]) || '' });
  });
  if (state.csvData) {
    state.pendingMappings.forEach((m) => {
      if (!state.csvData.columns.includes(m.column) && state.csvData.columns.length) m.column = state.csvData.columns[0];
    });
  }
  refreshColumnDropdowns();
  renderMappingsList();
  showToast('Preset loaded — assign columns if needed');
}

function verifyMappingsOnPage() {
  const fieldKeys = (state.pendingMappings || []).map((m) => m.fieldKey).filter(Boolean);
  if (!fieldKeys.length) {
    showToast('No mappings to verify — add mappings first', 'error');
    return;
  }
  sendToActiveTab({ type: 'VERIFY_MAPPINGS', fieldKeys }, (response) => {
    if (chrome.runtime.lastError) {
      showToast('Open a DocuSign tab and try again', 'error');
      return;
    }
    if (!response || !response.success) {
      showToast('Verification failed', 'error');
      return;
    }
    const missing = response.missing || [];
    if (missing.length === 0) {
      showToast('All ' + fieldKeys.length + ' mapping(s) found on page', 'success');
    } else {
      const labels = state.pendingMappings.filter((m) => missing.indexOf(m.fieldKey) >= 0).map((m) => m.fieldLabel || m.fieldKey);
      showToast(missing.length + ' not on page: ' + labels.slice(0, 3).join(', ') + (labels.length > 3 ? '…' : ''), 'error');
    }
    state.lastVerifyResult = { missing: response.missing || [], found: response.found || [] };
    refreshHealthCheck();
  });
}

function scanPageForFields() {
  if (!state.csvData || !state.csvData.columns.length) {
    showToast('Load a CSV/Excel file first (Data tab)', 'error');
    return;
  }

  const btn = $('scanFieldsBtn');
  btn.disabled = true;
  btn.textContent = 'Scanning…';

  sendToActiveTab({ type: 'SCAN_FIELDS' }, (response) => {
    btn.disabled = false;
    btn.textContent = '🔍 Scan page for fields';

    if (chrome.runtime.lastError) {
      showToast('Open a DocuSign tab and try again', 'error');
      return;
    }
    if (!response || !response.success || !response.fields || !response.fields.length) {
      showToast('No fields found on this page', 'error');
      return;
    }

    const columns = state.csvData.columns;
    const mappings = [];
    response.fields.forEach((f) => {
      const label = f.fieldLabel || f.fieldKey || '';
      let bestCol = '';
      let bestScore = 0.25;
      columns.forEach((col) => {
        const score = scoreColumnMatch(label, col);
        if (score > bestScore) { bestScore = score; bestCol = col; }
      });
      if (!bestCol && (f.fieldKey || '').toString().length > 2) {
        columns.forEach((col) => {
          const score = scoreColumnMatch((f.fieldKey || '').toString(), col);
          if (score > bestScore) { bestScore = score; bestCol = col; }
        });
      }
      if (bestCol) {
        mappings.push({
          fieldKey: f.fieldKey,
          fieldLabel: f.fieldLabel || f.fieldKey,
          column: bestCol,
          confidence: bestScore,
        });
      }
    });

    state.pendingMappings = mappings;
    renderMappingsList();

    const total = response.fields.length;
    const matched = mappings.length;
    showToast(
      `Scanned ${total} field(s), auto-mapped ${matched}. Add the rest with Start Capturing.`,
      matched ? 'success' : ''
    );
  });
}

function toggleCaptureMode() {
  state.captureMode = !state.captureMode;
  const btn = $('captureModeBtn');

  if (state.captureMode) {
    btn.textContent = '⏹ Stop Capturing';
    btn.style.borderColor = 'var(--accent)';
    btn.style.color = 'var(--accent)';
    $('mappingBanner').classList.add('active');
    sendToActiveTab({ type: 'CAPTURE_MODE_ON' });
    showToast('Click any field in DocuSign');
  } else {
    btn.textContent = '🎯 Start Capturing Fields';
    btn.style.borderColor = '';
    btn.style.color = '';
    $('mappingBanner').classList.remove('active');
    $('pendingCapture').style.display = 'none';
    sendToActiveTab({ type: 'CAPTURE_MODE_OFF' });
  }
}

function confirmMapping() {
  const colSelect = $('columnSelect');
  if (!state.pendingField || !colSelect.value) {
    showToast('Select a CSV column first', 'error');
    return;
  }

  const defInput = $('mappingDefaultInput');
  const transSelect = $('mappingTransformSelect');
  const newMapping = {
    fieldKey: state.pendingField.fieldKey,
    fieldLabel: state.pendingField.fieldLabel,
    column: colSelect.value,
  };
  if (defInput && defInput.value.trim()) newMapping.default = defInput.value.trim();
  if (transSelect && transSelect.value) newMapping.transform = transSelect.value;

  const existing = state.pendingMappings.findIndex(
    (m) => m.fieldKey === state.pendingField.fieldKey
  );
  if (existing >= 0) {
    state.pendingMappings[existing] = newMapping;
    showToast('Mapping updated (duplicate field replaced)', 'success');
  } else {
    state.pendingMappings.push(newMapping);
    showToast('Mapping added ✓', 'success');
  }

  state.pendingField = null;
  $('pendingCapture').style.display = 'none';
  if (defInput) defInput.value = '';
  if (transSelect) transSelect.value = '';
  renderMappingsList();
}

function cancelCapture() {
  state.pendingField = null;
  $('pendingCapture').style.display = 'none';
}

function renderMappingsList() {
  const list = $('mappingsList');
  $('mappingCount').textContent = `(${state.pendingMappings.length})`;

  if (state.pendingMappings.length === 0) {
    list.innerHTML =
      '<div class="empty-state">' +
      '<div class="empty-icon">↔️</div>' +
      '<div class="empty-text">No mappings yet.<br>Capture fields to build your template.</div>' +
      '</div>';
    return;
  }

  const sampleForColumn = (col) => {
    if (!state.csvData || !state.csvData.rows.length) return '';
    const vals = state.csvData.rows.slice(0, 3).map((r) => String(r[col] ?? '').slice(0, 20));
    return vals.join(', ');
  };

  list.innerHTML = state.pendingMappings
    .map((m, i) => {
      const defTag = m.default ? ' [def:' + escapeHtml(m.default) + ']' : '';
      const transTag = m.transform ? ' [' + escapeHtml(m.transform) + ']' : '';
      const colTitle = escapeHtml(m.fieldKey) + (m.column ? ' → sample: ' + escapeHtml(sampleForColumn(m.column)) : '');
      return `
        <div class="mapping-row" style="grid-template-columns:1fr 24px 1fr 28px 24px 24px;">
          <div class="field-tag" title="${escapeHtml(m.fieldKey)}">${escapeHtml(m.fieldLabel || m.fieldKey)}</div>
          <div class="mapping-arrow">→</div>
          <div class="field-tag" style="color:var(--accent);" title="${colTitle}">${escapeHtml(m.column)}${m.confidence != null ? confidenceBadge(m.confidence) : ''}${defTag}${transTag}</div>
          <button class="remove-btn" data-index="${i}" title="Remove">×</button>
          <button class="remove-btn" data-move-up="${i}" title="Move up" ${i === 0 ? 'disabled' : ''}>↑</button>
          <button class="remove-btn" data-move-down="${i}" title="Move down" ${i === state.pendingMappings.length - 1 ? 'disabled' : ''}>↓</button>
        </div>`;
    })
    .join('');

  // Remove buttons
  list.querySelectorAll('.remove-btn[data-index]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.pendingMappings.splice(parseInt(btn.dataset.index), 1);
      renderMappingsList();
    });
  });

  // Move up
  list.querySelectorAll('[data-move-up]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const i = parseInt(btn.dataset.moveUp, 10);
      if (i > 0) {
        [state.pendingMappings[i - 1], state.pendingMappings[i]] =
          [state.pendingMappings[i], state.pendingMappings[i - 1]];
        renderMappingsList();
      }
    });
  });

  // Move down
  list.querySelectorAll('[data-move-down]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const i = parseInt(btn.dataset.moveDown, 10);
      if (i < state.pendingMappings.length - 1) {
        [state.pendingMappings[i], state.pendingMappings[i + 1]] =
          [state.pendingMappings[i + 1], state.pendingMappings[i]];
        renderMappingsList();
      }
    });
  });
}


// ═════════════════════════════════════════════════════════════════════
// 10. FILL TAB (autofill & history)
// ═════════════════════════════════════════════════════════════════════

let filterRecordsDebounce;

function setupFillTab() {
  $('recordSearch').addEventListener('input', () => {
    clearTimeout(filterRecordsDebounce);
    filterRecordsDebounce = setTimeout(filterRecords, 120);
  });

  $('searchColumnSelect').addEventListener('change', () => {
    const col = $('searchColumnSelect').value;
    $('recordSearch').placeholder = col ? `Search in ${col}…` : 'Search by name, ID…';
    clearTimeout(filterRecordsDebounce);
    filterRecords();
  });

  $('clearFillMarkersBtn').addEventListener('click', () => {
    state.filledRecordIndices = [];
    saveToStorage();
    refreshRecords();
    showToast('Fill markers cleared');
  });

  $('recordSelect').addEventListener('change', (e) => {
    state.selectedRecord = e.target.value !== '' ? parseInt(e.target.value, 10) : null;
    updateFillButton();
    updatePreviewAndWarnings();
    saveToStorage();
  });

  $('templateSelect').addEventListener('change', () => {
    updateFillButton();
    updatePreviewAndWarnings();
    saveToStorage();
  });

  $('fillBtn').addEventListener('click', () => triggerFill(state.settings.batchFillMode));
  $('fillAndNextBtn').addEventListener('click', () => triggerFill(true));
  $('fillAllBatchBtn').addEventListener('click', fillAllBatch);
  $('fillNextBtn').addEventListener('click', fillNextRecord);
  $('retryFailedBtn').addEventListener('click', () => triggerFill(false, true));
  $('undoFillBtn').addEventListener('click', undoLastFill);
  $('duplicateRecordBtn').addEventListener('click', duplicateCurrentRecord);
  $('copyFillSummaryBtn').addEventListener('click', copyFillSummaryToClipboard);
  $('refreshConnectionBtn').addEventListener('click', refreshConnection);

  // Global Enter to fill when Fill tab is active (and not typing in an input)
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    const panel = document.querySelector('.panel.active');
    if (!panel || panel.id !== 'tab-fill') return;
    const t = e.target;
    const tag = t && t.tagName ? t.tagName.toLowerCase() : '';
    if (tag === 'input' || tag === 'select' || tag === 'textarea') return;
    const sel = $('recordSelect');
    if (sel && (t === sel || sel.contains(t))) return; // record list has its own Enter handler
    if ($('fillBtn').disabled) return;
    e.preventDefault();
    triggerFill(false);
  });

  // Keyboard navigation in record list
  $('recordSelect').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (!$('fillBtn').disabled) triggerFill(false);
      return;
    }

    const sel = $('recordSelect');
    const opts = sel.options;
    if (opts.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      const i = Math.min(sel.selectedIndex + 1, opts.length - 1);
      sel.selectedIndex = i;
      state.selectedRecord = parseInt(opts[i].value, 10);
      updateFillButton();
      updatePreviewAndWarnings();
      saveToStorage();
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      const i = Math.max(sel.selectedIndex - 1, 0);
      sel.selectedIndex = i;
      state.selectedRecord = parseInt(opts[i].value, 10);
      updateFillButton();
      updatePreviewAndWarnings();
      saveToStorage();
    }
  });
}

function updateFillButton() {
  const hasTemplate = !!$('templateSelect').value;
  const hasRecord = state.selectedRecord !== null;
  const ready = hasTemplate && hasRecord && state.isOnDocuSign;
  $('fillBtn').disabled = !ready;
  const fillAndNext = $('fillAndNextBtn');
  if (fillAndNext) fillAndNext.disabled = !ready;
  const fillAllBatch = $('fillAllBatchBtn');
  if (fillAllBatch) fillAllBatch.disabled = !ready || !state.csvData || state.csvData.rows.length === 0;
  refreshHealthCheck();
}

function updatePreviewAndWarnings() {
  const templateName = $('templateSelect').value;
  const previewWrap = $('previewBeforeFill');
  const previewList = $('previewList');
  const emptyWarn = $('emptyFieldsWarn');

  if (state.settings.showPreview === false && previewWrap) {
    previewWrap.style.display = 'none';
  }

  if (!templateName || !state.templates[templateName] || state.selectedRecord === null || !state.csvData) {
    if (previewWrap) previewWrap.style.display = 'none';
    if (emptyWarn) emptyWarn.style.display = 'none';
    return;
  }

  const record = state.csvData.rows[state.selectedRecord];
  const mappings = state.templates[templateName];
  const displayRecord = {};
  const emptyCols = [];

  mappings.forEach((m) => {
    if (m.conditionColumn && m.conditionValue) {
      if (String(record[m.conditionColumn] || '').trim() !== String(m.conditionValue).trim()) return;
    }
    const val = getRecordValue(record, m.column, m);
    displayRecord[m.column] = val;
    if (state.settings.warnEmpty && !val) emptyCols.push(m.column);
  });

  if (state.settings.showPreview && previewList) {
    previewWrap.style.display = 'block';
    previewList.innerHTML = mappings
      .map((m) => `<div>${escapeHtml(m.fieldLabel || m.fieldKey)} → ${escapeHtml(String(displayRecord[m.column] || '').slice(0, 40))}</div>`)
      .join('');
  }

  if (emptyWarn) {
    if (emptyCols.length && state.settings.warnEmpty) {
      emptyWarn.textContent = 'Empty for this record: ' + emptyCols.join(', ');
      emptyWarn.style.display = 'block';
    } else {
      emptyWarn.style.display = 'none';
    }
  }
}

// ── Fill execution ──────────────────────────────────────

function triggerFill(fillAndNext, retryFailedOnly, onComplete) {
  const templateName = $('templateSelect').value;
  if (!templateName || !state.templates[templateName]) {
    showToast('Select a template first', 'error');
    if (onComplete) onComplete();
    return;
  }
  if (state.selectedRecord === null || !state.csvData) {
    showToast('Select a record first', 'error');
    if (onComplete) onComplete();
    return;
  }

  const rawRecord = state.csvData.rows[state.selectedRecord];
  const mappings = state.templates[templateName];
  const displayRecord = {};
  mappings.forEach((m) => {
    if (m.conditionColumn && m.conditionValue) {
      if (String(rawRecord[m.conditionColumn] || '').trim() !== String(m.conditionValue).trim()) return;
    }
    displayRecord[m.column] = getRecordValue(rawRecord, m.column, m);
  });

  const failedKeysOnly =
    retryFailedOnly && state.lastFillResult?.failedKeys?.length
      ? state.lastFillResult.failedKeys
      : null;

  $('fillBtn').disabled = true;
  $('fillAndNextBtn').disabled = true;
  const fillAllBatchBtn = $('fillAllBatchBtn');
  if (fillAllBatchBtn) fillAllBatchBtn.disabled = true;
  $('fillBtn').innerHTML = '<div class="spinner"></div> Filling…';

  sendToActiveTab({ type: 'FILL_FIELDS', mappings, record: displayRecord, failedKeysOnly }, (response) => {
    $('fillBtn').disabled = false;
    $('fillBtn').textContent = '⚡ Autofill Document';
    if (fillAllBatchBtn) fillAllBatchBtn.disabled = !(state.templates[templateName] && state.selectedRecord !== null && state.csvData && state.isOnDocuSign);
    updateFillButton();

    const errEl = $('fillErrors');
    const summaryEl = $('fillSummary');
    const retryBtn = $('retryFailedBtn');
    const staleHint = $('staleTabHint');
    const fillTabHint = $('fillTabHint');
    const copySummaryBtn = $('copyFillSummaryBtn');

    if (errEl) errEl.style.display = 'none';
    if (summaryEl) summaryEl.style.display = 'none';
    if (copySummaryBtn) copySummaryBtn.style.display = 'none';
    if (retryBtn) retryBtn.style.display = 'none';
    if (staleHint) staleHint.style.display = 'none';
    if (fillTabHint && state.isOnDocuSign) fillTabHint.style.display = 'none';

    if (chrome.runtime.lastError) {
      showToast('Reload the DocuSign tab and try again', 'error');
      if (staleHint) staleHint.style.display = 'block';
      if (fillTabHint) fillTabHint.style.display = 'block';
      if (onComplete) onComplete();
      return;
    }

    // Store result for undo/retry and copy summary
    state.lastFillResult = response
      ? { count: response.count, errors: response.errors || [], failedKeys: response.failedKeys || [] }
      : null;
    state.lastRestoreData =
      response?.previousValues?.length ? response.previousValues : null;
    state.lastFillSummary = response && state.csvData && state.selectedRecord != null
      ? { templateName, recordIndex: state.selectedRecord, count: response.count, recordLabel: String(state.csvData.rows[state.selectedRecord][state.csvData.columns[0]] || 'Record ' + (state.selectedRecord + 1)) }
      : null;

    const undoBtn = $('undoFillBtn');
    if (undoBtn) undoBtn.style.display = state.lastRestoreData ? 'block' : 'none';

    // Track in fill history
    if (response && response.success && response.count > 0) {
      state.fillHistory.unshift({ recordIndex: state.selectedRecord, templateName, timestamp: Date.now() });
      if (state.fillHistory.length > FILL_HISTORY_MAX) state.fillHistory.pop();
      if (!state.filledRecordIndices.includes(state.selectedRecord)) {
        state.filledRecordIndices.push(state.selectedRecord);
      }
      saveToStorage();
      refreshFillHistory();
      refreshRecords();
    }

    // Display results
    if (response && response.failedKeys && response.failedKeys.length > 0) {
      showToast(response.count > 0
        ? `${response.count} filled, ${response.failedKeys.length} field(s) not found — check Mapping`
        : `${response.failedKeys.length} field(s) not found — re-capture in Mapping tab`, 'error');
    } else if (response && response.success) {
      showToast(`Filled ${response.count} field(s)`, 'success');
    }
    if (response && response.success) {
      renderFillSummary(summaryEl, errEl, retryBtn, response);
      if (copySummaryBtn && state.lastFillSummary) copySummaryBtn.style.display = 'block';
      if (fillAndNext || state.settings.batchFillMode) {
        if (state.settings.switchTabAfterFill && fillAndNext) {
          chrome.tabs.query({ currentWindow: true }, (tabs) => {
            const current = tabs.find((t) => t.active);
            if (current) {
              const idx = tabs.findIndex((t) => t.id === current.id);
              const next = tabs[idx + 1];
              if (next) chrome.tabs.update(next.id, { active: true });
            }
          });
        }
        fillNextRecord();
      }
    } else if (!response || !response.success) {
      if (!response || !response.failedKeys || response.failedKeys.length === 0) {
        showToast(response?.error || 'Fill failed — check console', 'error');
      }
      renderFillErrors(errEl, retryBtn, response);
    }
    if (onComplete) onComplete();
  });
}

function renderFillSummary(summaryEl, errEl, retryBtn, response) {
  if (!summaryEl) return;

  summaryEl.style.display = 'block';
  let html = '<strong>Filled ' + response.count + ' field(s)</strong>';

  if (response.errors && response.errors.length > 0) {
    const errText = escapeHtml(response.errors.slice(0, 3).join('; '));
    const more = response.errors.length > 3 ? ' (+' + (response.errors.length - 3) + ' more)' : '';
    html += '. <span style="color:var(--warning);">Failed: ' + errText + more + '</span>';
    if (retryBtn && response.failedKeys?.length) retryBtn.style.display = 'block';
  }

  html += ' <a href="#" id="fillSummaryGotoMapping" style="color:var(--accent);">Re-capture in Mapping</a>';
  summaryEl.innerHTML = html;

  const link = document.getElementById('fillSummaryGotoMapping');
  if (link) link.addEventListener('click', (e) => {
    e.preventDefault();
    document.querySelector('.tab[data-tab="mapping"]').click();
  });

  if (errEl && response.errors && response.errors.length > 0) {
    const errText = escapeHtml(response.errors.slice(0, 2).join('; '));
    const more = response.errors.length > 2 ? ' (+' + (response.errors.length - 2) + ' more)' : '';
    errEl.innerHTML =
      '<span style="color:var(--warning);">Some fields not found.</span> ' +
      errText + more +
      ' <a href="#" id="fillErrorsGotoMapping" style="color:var(--accent);">Re-capture in Mapping tab</a>';
    errEl.style.display = 'block';
    const errLink = document.getElementById('fillErrorsGotoMapping');
    if (errLink) errLink.addEventListener('click', (e) => {
      e.preventDefault();
      document.querySelector('.tab[data-tab="mapping"]').click();
    });
  }
}

function renderFillErrors(errEl, retryBtn, response) {
  if (!response?.errors?.length || !errEl) return;

  const errText = escapeHtml(response.errors.slice(0, 3).join('; '));
  const more = response.errors.length > 3 ? ' (+' + (response.errors.length - 3) + ' more)' : '';
  errEl.innerHTML =
    '<span style="color:var(--danger);">' + errText + more + '</span>. ' +
    '<a href="#" id="fillErrorsGotoMapping2" style="color:var(--accent);">Re-capture in Mapping tab</a>';
  errEl.style.display = 'block';

  const link = document.getElementById('fillErrorsGotoMapping2');
  if (link) link.addEventListener('click', (e) => {
    e.preventDefault();
    document.querySelector('.tab[data-tab="mapping"]').click();
  });

  if (response.failedKeys?.length && retryBtn) retryBtn.style.display = 'block';
}

// ── Undo & navigation ──────────────────────────────────

function undoLastFill() {
  if (!state.lastRestoreData || state.lastRestoreData.length === 0) {
    showToast('Nothing to undo', 'error');
    return;
  }

  sendToActiveTab({ type: 'RESTORE_FIELDS', restores: state.lastRestoreData }, (response) => {
    if (response && response.success) {
      showToast('Reverted ' + response.count + ' field(s)', 'success');
      state.lastRestoreData = null;
      $('undoFillBtn').style.display = 'none';
    } else if (chrome.runtime.lastError) {
      showToast('Reload the DocuSign tab and try again', 'error');
    }
  });
}

function refreshConnection() {
  showToast('Checking…');
  checkDocuSignTab();
  setTimeout(() => {
    showToast(state.isOnDocuSign && state.contentScriptReady ? 'Ready' : 'Reload the DocuSign tab', state.isOnDocuSign && state.contentScriptReady ? 'success' : 'error');
  }, 500);
}

function copyFillSummaryToClipboard() {
  if (!state.lastFillSummary) return;
  const s = state.lastFillSummary;
  const text = `Filled ${s.count} field(s). Record: ${s.recordLabel}. Template: ${s.templateName}.`;
  navigator.clipboard.writeText(text).then(() => showToast('Summary copied'), () => showToast('Copy failed', 'error'));
}

function fillAllBatch() {
  const indices = state.filteredIndices && state.filteredIndices.length > 0
    ? state.filteredIndices.slice()
    : state.csvData ? state.csvData.rows.map((_, i) => i) : [];
  if (indices.length === 0) {
    showToast('No records to fill', 'error');
    return;
  }
  let batchIndex = 0;
  const total = indices.length;
  const fillBtn = $('fillBtn');
  const fillAllBatchBtn = $('fillAllBatchBtn');

  function doOne() {
    if (batchIndex >= indices.length) {
      fillBtn.textContent = '⚡ Autofill Document';
      updateFillButton();
      showToast('Batch complete: ' + total + ' record(s)', 'success');
      return;
    }
    state.selectedRecord = indices[batchIndex];
    refreshRecords();
    updateFillButton();
    fillBtn.innerHTML = '<div class="spinner"></div> Filling ' + (batchIndex + 1) + ' of ' + total + '…';
    fillBtn.disabled = true;
    fillAllBatchBtn.disabled = true;
    triggerFill(false, false, () => {
      batchIndex++;
      setTimeout(doOne, 2000);
    });
  }
  doOne();
}

function duplicateCurrentRecord() {
  if (state.selectedRecord === null || !state.csvData) {
    showToast('Select a record first (Fill tab)', 'error');
    return;
  }
  const row = state.csvData.rows[state.selectedRecord];
  const copy = {};
  state.csvData.columns.forEach((col) => { copy[col] = row[col]; });
  state.csvData.rows.push(copy);
  state.selectedRecord = state.csvData.rows.length - 1;
  state.recordPage = Math.floor(state.selectedRecord / RECORD_PAGE_SIZE);
  saveToStorage();
  refreshRecords();
  updateFillButton();
  updatePreviewAndWarnings();
  showToast('Record duplicated');
}

function fillNextRecord() {
  if (!state.csvData || state.csvData.rows.length === 0) return;
  const next = state.selectedRecord === null
    ? 0
    : Math.min(state.selectedRecord + 1, state.csvData.rows.length - 1);
  state.selectedRecord = next;
  state.recordPage = Math.floor(next / RECORD_PAGE_SIZE);
  refreshRecords();
  updateFillButton();
  updatePreviewAndWarnings();
  saveToStorage();
}

function filterRecords() {
  const query = $('recordSearch').value.toLowerCase().trim();
  if (!query || !state.csvData) {
    state.filteredIndices = null;
    state.recordPage = 0;
    refreshRecords();
    return;
  }

  const rows = state.csvData.rows;
  const searchCol = $('searchColumnSelect')?.value || '';
  const cols = searchCol ? [searchCol] : state.csvData.columns;
  const matches = [];
  for (let i = 0; i < rows.length; i++) {
    for (let c = 0; c < cols.length; c++) {
      if (String(rows[i][cols[c]] || '').toLowerCase().includes(query)) {
        matches.push(i);
        break;
      }
    }
  }
  state.filteredIndices = matches;
  state.recordPage = 0;
  refreshRecords();
}

function refreshFillHistory() {
  const wrap = $('fillHistoryWrap');
  const list = $('fillHistoryList');
  if (!wrap || !list) return;

  if (state.fillHistory.length === 0) { wrap.style.display = 'none'; return; }

  wrap.style.display = 'block';
  const fmt = (ts) => {
    const d = new Date(ts);
    return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  };
  list.innerHTML = state.fillHistory
    .slice(0, 5)
    .map((h) => `<div>${h.templateName} · Record ${(h.recordIndex || 0) + 1} · ${fmt(h.timestamp)}</div>`)
    .join('');
}


// ═════════════════════════════════════════════════════════════════════
// 11. UI REFRESH & RENDERING
// ═════════════════════════════════════════════════════════════════════

function refreshAll() {
  refreshCSVStatus();
  refreshSheetPicker();
  refreshRecords();
  refreshFillTemplates();
  refreshSavedTemplates();
  refreshColumnDropdowns();
  refreshSearchColumnDropdown();
  updatePreviewAndWarnings();
  updateFillButton();
  refreshFillHistory();
  refreshHealthCheck();
}

function refreshHealthCheck() {
  const docuSignEl = $('healthDocuSignVal');
  const scriptEl = $('healthScriptVal');
  const dataEl = $('healthDataVal');
  const templateEl = $('healthTemplateVal');
  const mappingsEl = $('healthMappingsVal');
  const recordEl = $('healthRecordVal');
  const emptyEl = $('healthEmptyVal');
  const columnsEl = $('healthColumnsVal');
  const readyEl = $('healthReadyVal');
  if (!docuSignEl) return;

  if (state.isOnDocuSign) {
    docuSignEl.textContent = 'Yes';
    docuSignEl.className = 'health-ok';
  } else {
    docuSignEl.textContent = 'No';
    docuSignEl.className = 'health-no';
  }

  if (state.isOnDocuSign && state.contentScriptReady) {
    scriptEl.textContent = 'Ready';
    scriptEl.className = 'health-ok';
  } else if (state.isOnDocuSign) {
    scriptEl.textContent = 'Not loaded';
    scriptEl.className = 'health-no';
  } else {
    scriptEl.textContent = '—';
    scriptEl.className = 'health-no';
  }

  if (state.csvData && state.csvData.rows.length > 0) {
    dataEl.textContent = 'Loaded (' + state.csvData.rows.length + ' rows)';
    dataEl.className = 'health-ok';
  } else if (state.csvData) {
    dataEl.textContent = 'Loaded (0 rows)';
    dataEl.className = 'health-no';
  } else {
    dataEl.textContent = 'No data';
    dataEl.className = 'health-no';
  }

  const templateName = $('templateSelect')?.value || '';
  let mappingsCount = 0;
  if (templateName && state.templates[templateName]) {
    templateEl.textContent = 'Selected';
    templateEl.className = 'health-ok';
    mappingsCount = state.templates[templateName].length;
    if (mappingsEl) {
      mappingsEl.textContent = mappingsCount + ' (for this template)';
      mappingsEl.className = mappingsCount > 0 ? 'health-ok' : 'health-no';
    }
  } else {
    templateEl.textContent = 'None';
    templateEl.className = 'health-no';
    if (mappingsEl) {
      mappingsEl.textContent = '—';
      mappingsEl.className = 'health-no';
    }
  }

  if (state.selectedRecord !== null && state.csvData) {
    recordEl.textContent = 'Selected';
    recordEl.className = 'health-ok';
  } else {
    recordEl.textContent = 'None';
    recordEl.className = 'health-no';
  }

  // Empty fields for this record (mapped columns that are empty)
  let emptyCount = 0;
  if (emptyEl && templateName && state.templates[templateName] && state.selectedRecord !== null && state.csvData) {
    const record = state.csvData.rows[state.selectedRecord];
    const mappings = state.templates[templateName];
    mappings.forEach((m) => {
      const val = getRecordValue(record, m.column, m);
      if (val === undefined || val === null || String(val).trim() === '') emptyCount++;
    });
    emptyEl.textContent = emptyCount === 0 ? 'None' : emptyCount + ' empty';
    emptyEl.className = emptyCount === 0 ? 'health-ok' : 'health-no';
  } else if (emptyEl) {
    emptyEl.textContent = '—';
    emptyEl.className = 'health-no';
  }

  // Template columns present in CSV
  if (columnsEl && templateName && state.templates[templateName] && state.csvData) {
    const cols = state.csvData.columns;
    const needed = [];
    state.templates[templateName].forEach((m) => {
      if (m.column && needed.indexOf(m.column) < 0) needed.push(m.column);
    });
    const missing = needed.filter((c) => cols.indexOf(c) < 0);
    if (missing.length === 0) {
      columnsEl.textContent = 'All present';
      columnsEl.className = 'health-ok';
    } else {
      columnsEl.textContent = 'Missing: ' + missing.slice(0, 3).join(', ') + (missing.length > 3 ? '…' : '');
      columnsEl.className = 'health-no';
    }
  } else if (columnsEl) {
    columnsEl.textContent = '—';
    columnsEl.className = 'health-no';
  }

  // DocuSign tabs open
  const tabsEl = $('healthTabsVal');
  if (tabsEl) {
    if (state.isOnDocuSign) {
      tabsEl.textContent = state.docuSignTabCount + ' open (using active)';
      tabsEl.className = 'health-ok';
    } else {
      tabsEl.textContent = state.docuSignTabCount > 0 ? state.docuSignTabCount + ' open' : '—';
      tabsEl.className = 'health-no';
    }
  }

  // Last fill
  const lastFillEl = $('healthLastFillVal');
  if (lastFillEl) {
    if (state.lastFillResult) {
      if (state.lastFillResult.count != null && state.lastFillResult.count > 0) {
        lastFillEl.textContent = state.lastFillResult.count + ' field(s)';
        lastFillEl.className = 'health-ok';
      } else if (state.lastFillResult.failedKeys && state.lastFillResult.failedKeys.length > 0) {
        lastFillEl.textContent = 'Failed (' + state.lastFillResult.failedKeys.length + ' missing)';
        lastFillEl.className = 'health-no';
      } else {
        lastFillEl.textContent = 'Failed';
        lastFillEl.className = 'health-no';
      }
    } else {
      lastFillEl.textContent = '—';
      lastFillEl.className = 'health-no';
    }
  }

  // Last verify (from Mapping tab)
  const verifyEl = $('healthVerifyVal');
  if (verifyEl) {
    if (state.lastVerifyResult) {
      const missing = state.lastVerifyResult.missing || [];
      if (missing.length === 0) {
        verifyEl.textContent = 'All verified';
        verifyEl.className = 'health-ok';
      } else {
        verifyEl.textContent = missing.length + ' missing on page';
        verifyEl.className = 'health-no';
      }
    } else {
      verifyEl.textContent = '—';
      verifyEl.className = 'health-no';
    }
  }

  // CSV columns not in template (unmapped columns)
  const unmappedEl = $('healthUnmappedVal');
  if (unmappedEl && state.csvData && templateName && state.templates[templateName]) {
    const cols = state.csvData.columns;
    const mapped = [];
    state.templates[templateName].forEach((m) => {
      if (m.column && mapped.indexOf(m.column) < 0) mapped.push(m.column);
    });
    const notInTemplate = cols.filter((c) => mapped.indexOf(c) < 0);
    if (notInTemplate.length === 0) {
      unmappedEl.textContent = 'All mapped';
      unmappedEl.className = 'health-ok';
    } else {
      unmappedEl.textContent = notInTemplate.length + ' not in template';
      unmappedEl.className = 'health-no';
    }
  } else if (unmappedEl) {
    unmappedEl.textContent = '—';
    unmappedEl.className = 'health-no';
  }

  // Extension version
  const versionEl = $('healthVersionVal');
  if (versionEl) {
    try {
      const v = chrome.runtime.getManifest().version;
      versionEl.textContent = 'v' + (v || '—');
      versionEl.className = 'health-ok';
    } catch (e) {
      versionEl.textContent = '—';
      versionEl.className = 'health-no';
    }
  }

  // Ready to fill: all conditions
  const ready = readyEl && state.isOnDocuSign && state.contentScriptReady &&
    state.csvData && state.csvData.rows.length > 0 &&
    templateName && state.templates[templateName] && state.templates[templateName].length > 0 &&
    state.selectedRecord !== null;
  if (readyEl) {
    readyEl.textContent = ready ? 'Yes' : 'No';
    readyEl.className = ready ? 'health-ok' : 'health-no';
  }
}

function refreshCSVStatus() {
  const dot = $('csvDot');
  if (state.csvData) {
    dot.classList.add('loaded');
    $('csvName').textContent = state.csvData.name;
    $('csvCount').textContent = `${state.csvData.rows.length} rows`;
    $('fileInfo').style.display = 'block';
    $('fileInfoName').textContent = state.csvData.name;
    $('fileInfoRows').textContent = state.csvData.rows.length;
    $('fileInfoCols').textContent = state.csvData.columns.length;

    $('columnsList').innerHTML = state.csvData.columns
      .map((col) =>
        '<span style="padding:3px 7px;background:var(--surface2);border:1px solid var(--border);border-radius:4px;font-size:10px;color:var(--text-muted);">' +
        escapeHtml(col) + '</span>'
      )
      .join('');

    const cols = state.csvData.columns;
    const rows = state.csvData.rows.slice(0, 3);
    $('dataPreview').innerHTML =
      '<table style="border-collapse:collapse;width:100%;min-width:300px;">' +
      '<thead><tr>' +
      cols.map((c) => '<th style="border:1px solid var(--border);padding:4px 6px;text-align:left;color:var(--text-dim);white-space:nowrap;">' + escapeHtml(c) + '</th>').join('') +
      '</tr></thead>' +
      '<tbody>' +
      rows.map((r) => '<tr>' +
        cols.map((c) => '<td style="border:1px solid var(--border);padding:3px 6px;color:var(--text-muted);white-space:nowrap;max-width:80px;overflow:hidden;text-overflow:ellipsis;">' + escapeHtml(r[c] ?? '') + '</td>').join('') +
        '</tr>'
      ).join('') +
      '</tbody></table>';
  } else {
    dot.classList.remove('loaded');
    $('csvName').textContent = 'No file loaded';
    $('csvCount').textContent = '';
    $('fileInfo').style.display = 'none';
  }
}

function refreshSheetPicker() {
  const picker = $('sheetPicker');
  if (!picker) return;

  if (state.pendingXlsxSheets && state.pendingXlsxSheets.length > 0) {
    picker.style.display = 'block';
    const sel = $('sheetSelect');
    sel.innerHTML = state.pendingXlsxSheets
      .map((name, i) => `<option value="${i}">${escapeHtml(name)}</option>`)
      .join('');
  } else {
    picker.style.display = 'none';
  }
}

function refreshRecords() {
  const sel = $('recordSelect');
  const paginationEl = $('recordPagination');
  const progressEl = $('recordProgress');

  if (!state.csvData) {
    sel.innerHTML = '<option value="">Load a CSV first</option>';
    if (paginationEl) paginationEl.style.display = 'none';
    if (progressEl) progressEl.textContent = '';
    return;
  }

  const rows = state.csvData.rows;
  const cols = state.csvData.columns;
  const displayCol = cols.find((c) => /name|provider|label|id|first/i.test(c)) || cols[0];

  // Use filtered indices if a search is active, otherwise all indices
  const isFiltered = state.filteredIndices !== null;
  const indices = isFiltered ? state.filteredIndices : rows.map((_, i) => i);
  const total = indices.length;

  if (isFiltered && total === 0) {
    sel.innerHTML = '<option value="">No matching records</option>';
    if (paginationEl) paginationEl.style.display = 'none';
    if (progressEl) progressEl.textContent = '(0 matches)';
    return;
  }

  const label = isFiltered ? 'matches' : 'records';

  const filled = state.filledRecordIndices;
  const optionText = (idx) => {
    const prefix = filled.includes(idx) ? '\u2713 ' : '';
    return prefix + escapeHtml(String(rows[idx][displayCol] || '').trim() || `Row ${idx + 1}`);
  };

  if (total <= RECORD_PAGE_SIZE) {
    // No pagination needed
    sel.innerHTML = indices
      .map((idx) => `<option value="${idx}">${optionText(idx)}</option>`)
      .join('');
    if (paginationEl) paginationEl.style.display = 'none';
  } else {
    // Paginated view
    const start = state.recordPage * RECORD_PAGE_SIZE;
    const end = Math.min(start + RECORD_PAGE_SIZE, total);
    const pageIndices = indices.slice(start, end);

    sel.innerHTML = pageIndices
      .map((idx) => `<option value="${idx}">${optionText(idx)}</option>`)
      .join('');

    if (paginationEl) {
      paginationEl.style.display = 'flex';
      paginationEl.innerHTML = `
        <span>${start + 1}–${end} of ${total} ${label}</span>
        <button type="button" class="btn btn-secondary btn-sm" id="recordPrevPage" ${state.recordPage === 0 ? 'disabled' : ''}>Prev</button>
        <button type="button" class="btn btn-secondary btn-sm" id="recordNextPage" ${end >= total ? 'disabled' : ''}>Next</button>
      `;
      const prevBtn = $('recordPrevPage');
      const nextBtn = $('recordNextPage');
      if (prevBtn) prevBtn.addEventListener('click', () => {
        state.recordPage = Math.max(0, state.recordPage - 1);
        refreshRecords();
        updateFillButton();
      });
      if (nextBtn) nextBtn.addEventListener('click', () => {
        state.recordPage = Math.min(Math.floor((total - 1) / RECORD_PAGE_SIZE), state.recordPage + 1);
        refreshRecords();
        updateFillButton();
      });
    }
  }

  if (progressEl) progressEl.textContent = total ? `(${total} ${label})` : '';

  const clearMarkersBtn = $('clearFillMarkersBtn');
  if (clearMarkersBtn) {
    clearMarkersBtn.style.display = filled.length > 0 ? 'block' : 'none';
  }

  // Restore selection
  if (state.selectedRecord !== null && state.selectedRecord >= 0 && state.selectedRecord < rows.length) {
    if (total > RECORD_PAGE_SIZE) {
      const start = state.recordPage * RECORD_PAGE_SIZE;
      const end = Math.min(start + RECORD_PAGE_SIZE, total);
      const pageIndices = indices.slice(start, end);
      if (pageIndices.includes(state.selectedRecord)) {
        sel.value = String(state.selectedRecord);
      }
      progressEl.textContent = `(Record ${state.selectedRecord + 1} of ${rows.length}${isFiltered ? ', ' + total + ' matches' : ''})`;
    } else {
      sel.value = String(state.selectedRecord);
      progressEl.textContent = `(Record ${state.selectedRecord + 1} of ${rows.length}${isFiltered ? ', ' + total + ' matches' : ''})`;
    }
  }
}

function refreshColumnDropdowns() {
  const sel = $('columnSelect');
  if (!state.csvData) {
    sel.innerHTML = '<option value="">— Load CSV first —</option>';
    return;
  }
  sel.innerHTML =
    '<option value="">— Select column —</option>' +
    state.csvData.columns
      .map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`)
      .join('');
}

function refreshSearchColumnDropdown() {
  const sel = $('searchColumnSelect');
  if (!sel) return;
  if (!state.csvData) {
    sel.innerHTML = '<option value="">All columns</option>';
    return;
  }
  sel.innerHTML =
    '<option value="">All columns</option>' +
    state.csvData.columns
      .map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`)
      .join('');
}


// ═════════════════════════════════════════════════════════════════════
// 12. DOCUSIGN DETECTION
// ═════════════════════════════════════════════════════════════════════

function checkDocuSignTab() {
  chrome.tabs.query({}, (allTabs) => {
    state.docuSignTabCount = (allTabs || []).filter((t) => t.url && /docusign\.(com|net)/.test(t.url)).length;

    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      const tab = tabs[0];

    if (tab && tab.url && /docusign\.(com|net)/.test(tab.url)) {
      state.isOnDocuSign = true;
      $('pageStatus').textContent = 'On DocuSign ✓';
      $('pageStatus').style.color = 'var(--success)';
      $('fillTabHint').style.display = 'none';

      try {
        chrome.tabs.sendMessage(tab.id, { type: 'PING' }, (response) => {
          state.contentScriptReady = !!response && response.ok;
          const staleHint = $('staleTabHint');
          if (staleHint) staleHint.style.display = state.contentScriptReady ? 'none' : 'block';
          updateFillButton();
          refreshHealthCheck();
        });
      } catch (e) {
        state.contentScriptReady = false;
        $('staleTabHint').style.display = 'block';
        updateFillButton();
        refreshHealthCheck();
      }
    } else {
      state.isOnDocuSign = false;
      state.contentScriptReady = false;
      $('pageStatus').textContent = 'Not on DocuSign — open a DocuSign tab';
      $('pageStatus').style.color = 'var(--text-dim)';
      $('fillTabHint').style.display = 'block';
      $('staleTabHint').style.display = 'none';
    }

    updateFillButton();
    refreshHealthCheck();
  });
  });
}


// ═════════════════════════════════════════════════════════════════════
// 13. MESSAGING
// ═════════════════════════════════════════════════════════════════════

function sendToActiveTab(message, callback) {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (!tabs[0]) {
      if (callback) callback({ error: 'No active tab' });
      return;
    }
    try {
      chrome.tabs.sendMessage(tabs[0].id, message, (response) => {
        if (callback) callback(response);
      });
    } catch (e) {
      if (callback) callback({ error: 'Extension not ready on this page' });
    }
  });
}

function listenForContentMessages() {
  chrome.runtime.onMessage.addListener((message) => {
    if (message.type === 'FIELD_CAPTURED') {
      state.pendingField = { fieldKey: message.fieldKey, fieldLabel: message.fieldLabel };
      $('capturedFieldName').textContent = message.fieldLabel || message.fieldKey;
      $('pendingCapture').style.display = 'block';
      refreshColumnDropdowns();
      const label = message.fieldLabel || message.fieldKey || '';
      const columns = state.csvData ? state.csvData.columns : [];
      let bestCol = '';
      let bestScore = 0.25;
      columns.forEach((col) => {
        const score = scoreColumnMatch(label, col);
        if (score > bestScore) { bestScore = score; bestCol = col; }
      });
      if (!bestCol && (message.fieldKey || '').toString().length > 2) {
        columns.forEach((col) => {
          const score = scoreColumnMatch((message.fieldKey || '').toString(), col);
          if (score > bestScore) { bestScore = score; bestCol = col; }
        });
      }
      const colSelect = $('columnSelect');
      if (colSelect && bestCol) colSelect.value = bestCol;
      document.querySelector('.tab[data-tab="mapping"]').click();
    }
  });
}
