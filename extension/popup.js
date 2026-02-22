// ═══════════════════════════════════════════════════════
//  DocuFill – popup.js
// ═══════════════════════════════════════════════════════

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
};

const RECORD_PAGE_SIZE = 50;

let state = {
  csvData: null,
  templates: {},
  activeTemplate: null,
  pendingMappings: [],
  captureMode: false,
  pendingField: null,
  selectedRecord: null,
  isOnDocuSign: false,
  pendingXlsxData: null,
  pendingXlsxSheets: null,
  settings: { ...DEFAULT_SETTINGS },
  onboardingDismissed: false,
  lastFillResult: null,
  lastRestoreData: null,
  fillHistory: [],
  recordPage: 0,
  contentScriptReady: false,
};

const FILL_HISTORY_MAX = 10;

const $ = (id) => document.getElementById(id);

function escapeHtml(str) {
  if (str == null) return '';
  const s = String(str);
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const STORAGE_QUOTA_WARN_BYTES = 2.5 * 1024 * 1024; // ~2.5 MB

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

// ── Storage ─────────────────────────────────────────────

async function loadFromStorage() {
  return new Promise((resolve) => {
    chrome.storage.local.get(['csvData', 'templates', 'lastTemplate', 'lastRecord', 'settings', 'onboardingDismissed', 'fillHistory'], (result) => {
      if (result.csvData) state.csvData = result.csvData;
      if (result.templates) state.templates = result.templates;
      if (result.settings && typeof result.settings === 'object') state.settings = { ...DEFAULT_SETTINGS, ...result.settings };
      if (result.onboardingDismissed) state.onboardingDismissed = true;
      if (Array.isArray(result.fillHistory)) state.fillHistory = result.fillHistory.slice(0, FILL_HISTORY_MAX);
      refreshAll();
      if (result.lastTemplate && state.templates[result.lastTemplate]) {
        $('templateSelect').value = result.lastTemplate;
        state.selectedRecord = null;
      }
      if (result.lastRecord != null && state.csvData && result.lastRecord >= 0 && result.lastRecord < state.csvData.rows.length) {
        state.selectedRecord = result.lastRecord;
        state.recordPage = Math.floor(result.lastRecord / RECORD_PAGE_SIZE);
        $('recordSelect').value = String(result.lastRecord);
      }
      updateFillButton();
      refreshSettingsUI();
      resolve();
    });
  });
}

function saveToStorage() {
  const payload = { csvData: state.csvData, templates: state.templates, settings: state.settings, onboardingDismissed: state.onboardingDismissed, fillHistory: state.fillHistory.slice(0, FILL_HISTORY_MAX) };
  const lastTemplate = $('templateSelect')?.value || '';
  if (lastTemplate) payload.lastTemplate = lastTemplate;
  if (state.csvData && state.selectedRecord != null) payload.lastRecord = state.selectedRecord;
  else payload.lastRecord = null;
  chrome.storage.local.set(payload);
}

// ── Tabs ────────────────────────────────────────────────

function setupTabs() {
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach((t) => { t.classList.remove('active'); t.setAttribute('aria-selected', 'false'); });
      document.querySelectorAll('.panel').forEach((p) => p.classList.remove('active'));
      tab.classList.add('active');
      tab.setAttribute('aria-selected', 'true');
      const panel = $('tab-' + tab.dataset.tab);
      if (panel) panel.classList.add('active');
    });
  });
}

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
  const themeEl = $('settingTheme');
  if (themeEl) themeEl.value = s.theme || 'dark';
  const trimEl = $('settingTrimWhitespace');
  if (trimEl) trimEl.checked = s.trimWhitespace !== false;
  const dateEl = $('settingDateFormat');
  if (dateEl) dateEl.value = s.dateFormat || '';
  const encEl = $('settingEncoding');
  if (encEl) encEl.value = s.encoding || 'utf-8';
  const autoEl = $('settingAutoLoadTemplate');
  if (autoEl) autoEl.checked = !!s.autoLoadTemplate;
  const rulesEl = $('settingUrlRules');
  if (rulesEl) rulesEl.value = s.urlRules || '';
  const wrap = $('urlRulesWrap');
  if (wrap) wrap.style.display = autoEl && autoEl.checked ? 'block' : 'none';
  const warnEl = $('settingWarnEmpty');
  if (warnEl) warnEl.checked = s.warnEmpty !== false;
  const previewEl = $('settingShowPreview');
  if (previewEl) previewEl.checked = s.showPreview !== false;
  const batchEl = $('settingBatchFillMode');
  if (batchEl) batchEl.checked = !!s.batchFillMode;
}

function setupSettingsTab() {
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
    };
    applyTheme(state.settings.theme);
    $('urlRulesWrap').style.display = state.settings.autoLoadTemplate ? 'block' : 'none';
    saveToStorage();
    showToast('Settings saved');
  }
  $('settingTheme').addEventListener('change', saveSettings);
  $('settingTrimWhitespace').addEventListener('change', saveSettings);
  $('settingDateFormat').addEventListener('change', saveSettings);
  $('settingEncoding').addEventListener('change', saveSettings);
  $('settingAutoLoadTemplate').addEventListener('change', saveSettings);
  $('settingUrlRules').addEventListener('change', saveSettings);
  $('settingWarnEmpty').addEventListener('change', saveSettings);
  $('settingShowPreview').addEventListener('change', saveSettings);
  $('settingBatchFillMode').addEventListener('change', saveSettings);
}

function setupHelpTab() {}

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

function tryApplyUrlTemplateRules() {
  if (!state.settings.autoLoadTemplate || !state.settings.urlRules) return;
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const url = tabs[0]?.url || '';
    const lines = state.settings.urlRules.split('\n').map(l => l.trim()).filter(Boolean);
    for (const line of lines) {
      const m = line.match(/^(\S+)\s+(.+)$/);
      if (!m) continue;
      let pattern = m[1];
      const templateName = m[2].trim();
      if (!pattern || !templateName) continue;
      if (pattern.startsWith('/') && pattern.endsWith('/')) pattern = pattern.slice(1, -1);
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
  if (state.settings.dateFormat && /date|dob|birth|year/i.test(column)) {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      const pad = n => String(n).padStart(2, '0');
      const y = d.getFullYear(), m = pad(d.getMonth() + 1), day = pad(d.getDate());
      if (state.settings.dateFormat === 'MM/DD/YYYY') val = `${m}/${day}/${y}`;
      else if (state.settings.dateFormat === 'DD/MM/YYYY') val = `${day}/${m}/${y}`;
      else if (state.settings.dateFormat === 'YYYY-MM-DD') val = `${y}-${m}-${day}`;
    }
  }
  return val;
}

// ── Data Tab ────────────────────────────────────────────

function setupDataTab() {
  const uploadZone = $('uploadZone');
  const fileInput = $('fileInput');

  uploadZone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => handleFileUpload(e.target.files[0]));

  uploadZone.addEventListener('dragover', (e) => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
  uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
  uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFileUpload(e.dataTransfer.files[0]);
  });

  $('clearFileBtn').addEventListener('click', () => {
    state.csvData = null;
    state.selectedRecord = null;
    saveToStorage();
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
  $('importTemplatesInput').addEventListener('change', (e) => { if (e.target.files[0]) importTemplates(e.target.files[0]); e.target.value = ''; });
  $('importMappingsCsvBtn').addEventListener('click', () => $('importMappingsCsvInput').click());
  $('importMappingsCsvInput').addEventListener('change', (e) => { if (e.target.files[0]) importMappingsCsv(e.target.files[0]); e.target.value = ''; });
  $('loadSampleCsvBtn').addEventListener('click', loadSampleCsv);
}

const SAMPLE_CSV = 'Name,Email,Company\nJane Doe,jane@example.com,Acme Inc\nJohn Smith,john@example.com,Widget Co';

function loadSampleCsv(silent) {
  try {
    const parsed = parseCSV(SAMPLE_CSV);
    state.csvData = { name: 'sample.csv', columns: parsed.columns, rows: parsed.rows };
    saveToStorage();
    refreshAll();
    if (!silent) showToast('Loaded sample CSV (' + parsed.rows.length + ' rows)', 'success');
  } catch (e) {
    if (!silent) showToast('Failed to load sample', 'error');
  }
}

async function importMappingsCsv(file) {
  try {
    const text = await file.text();
    const lines = text.trim().split(/\r?\n/).filter(Boolean);
    if (lines.length < 2) { showToast('CSV needs header + at least 1 row', 'error'); return; }
    const cols = splitCSVLine(lines[0]);
    const fieldKeyIdx = cols.findIndex(c => /fieldkey|field_key/i.test(c));
    const fieldLabelIdx = cols.findIndex(c => /fieldlabel|field_label/i.test(c));
    const columnIdx = cols.findIndex(c => /column/i.test(c));
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
    const keys = Object.keys(imported);
    let merged = 0;
    keys.forEach((k) => {
      if (Array.isArray(imported[k])) {
        state.templates[k] = imported[k];
        merged++;
      }
    });
    saveToStorage();
    refreshAll();
    showToast(merged ? `Imported ${merged} template(s)` : 'No valid templates in file', merged ? 'success' : 'error');
  } catch (err) {
    showToast('Import failed: ' + (err.message || 'invalid file'), 'error');
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
    const csvData = { name: state.pendingXlsxFileName || 'Sheet', columns, rows };
    const approxBytes = JSON.stringify(csvData).length;
    if (approxBytes > STORAGE_QUOTA_WARN_BYTES) {
      showToast('File is very large; storage may fail. Try a smaller file.', 'error');
      return;
    }
    state.csvData = csvData;
    state.pendingXlsxData = null;
    state.pendingXlsxSheets = null;
    state.pendingXlsxFileName = null;
    saveToStorage();
    refreshAll();
    showToast(rows.length ? `Loaded ${rows.length} records` : 'Columns loaded (0 rows)', rows.length ? 'success' : '');
  } catch (err) {
    showToast(err.message || 'Failed to load sheet', 'error');
  }
}

async function handleFileUpload(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  try {
    let rows, columns, duplicateColumns = null;
    if (ext === 'csv') {
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
    const csvData = { name: file.name, columns, rows };
    const approxBytes = JSON.stringify(csvData).length;
    if (approxBytes > STORAGE_QUOTA_WARN_BYTES) {
      showToast('File is very large; storage may fail. Try a smaller file.', 'error');
      return;
    }
    state.csvData = csvData;
    saveToStorage();
    refreshAll();
    showToast(rows.length ? `Loaded ${rows.length} records` : 'Columns loaded (0 rows)', rows.length ? 'success' : '');
  } catch (err) {
    console.error(err);
    showToast('Failed to parse file', 'error');
  }
}

function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 1) throw new Error('CSV is empty');
  const columns = splitCSVLine(lines[0]);
  const rows = lines.length < 2 ? [] : lines.slice(1).map((line) => {
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
    if (ch === '"') { inQuotes = !inQuotes; }
    else if (ch === ',' && !inQuotes) { result.push(current.trim()); current = ''; }
    else { current += ch; }
  }
  result.push(current.trim());
  return result;
}

function sendParseXLSX(dataArray, sheetIndex) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(
      { type: 'PARSE_XLSX', data: dataArray, sheetIndex: sheetIndex },
      (response) => {
        if (response && response.success) {
          if (response.needSheetChoice) resolve({ needSheetChoice: true, sheetNames: response.sheetNames });
          else resolve({ result: response.result });
        } else reject(new Error(response?.error || 'Parse failed'));
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

// ── Mapping Tab ─────────────────────────────────────────

function setupMappingTab() {
  $('scanFieldsBtn').addEventListener('click', scanPageForFields);
  $('captureModeBtn').addEventListener('click', toggleCaptureMode);
  $('confirmMappingBtn').addEventListener('click', confirmMapping);
  $('cancelCaptureBtn').addEventListener('click', cancelCapture);
  $('saveTemplateBtn').addEventListener('click', saveTemplate);
  $('newTemplateBtn').addEventListener('click', newTemplate);
  $('loadTemplateBtn').addEventListener('click', loadTemplateIntoEditor);
}

function normalizeForMatch(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim().replace(/[^a-z0-9\s]/g, '');
}

function scoreColumnMatch(fieldLabel, column) {
  if (!column) return 0;
  const a = normalizeForMatch(fieldLabel);
  const b = normalizeForMatch(column);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (fieldLabel.toLowerCase() === column.toLowerCase()) return 0.98;
  const aNorm = a.replace(/\s/g, '');
  const bNorm = b.replace(/\s/g, '');
  if (aNorm === bNorm) return 0.95;
  if (b.includes(a) || a.includes(b)) return 0.7;
  const aWords = a.split(/\s+/).filter(Boolean);
  const bWords = b.split(/\s+/).filter(Boolean);
  const overlap = aWords.filter(w => bWords.some(bw => bw.includes(w) || w.includes(bw))).length;
  if (overlap > 0) return 0.4 + (overlap / Math.max(aWords.length, bWords.length)) * 0.4;
  return 0;
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
      let bestScore = 0.3;
      columns.forEach((col) => {
        const score = scoreColumnMatch(label, col);
        if (score > bestScore) { bestScore = score; bestCol = col; }
      });
      if (bestCol) {
        mappings.push({ fieldKey: f.fieldKey, fieldLabel: f.fieldLabel || f.fieldKey, column: bestCol });
      }
    });
    state.pendingMappings = mappings;
    renderMappingsList();
    const total = response.fields.length;
    const matched = mappings.length;
    showToast(`Scanned ${total} field(s), auto-mapped ${matched}. Add the rest with Start Capturing.`, matched ? 'success' : '');
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
  const newMapping = { fieldKey: state.pendingField.fieldKey, fieldLabel: state.pendingField.fieldLabel, column: colSelect.value };
  if (defInput && defInput.value.trim()) newMapping.default = defInput.value.trim();
  if (transSelect && transSelect.value) newMapping.transform = transSelect.value;
  const existing = state.pendingMappings.findIndex(m => m.fieldKey === state.pendingField.fieldKey);
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

function renderMappingsList() {
  const list = $('mappingsList');
  $('mappingCount').textContent = `(${state.pendingMappings.length})`;
  if (state.pendingMappings.length === 0) {
    list.innerHTML = `<div class="empty-state"><div class="empty-icon">↔️</div><div class="empty-text">No mappings yet.<br>Capture fields to build your template.</div></div>`;
    return;
  }
  list.innerHTML = state.pendingMappings.map((m, i) => `
    <div class="mapping-row" style="grid-template-columns:1fr 24px 1fr 28px 24px 24px;">
      <div class="field-tag" title="${escapeHtml(m.fieldKey)}">${escapeHtml(m.fieldLabel || m.fieldKey)}</div>
      <div class="mapping-arrow">→</div>
      <div class="field-tag" style="color:var(--accent);">${escapeHtml(m.column)}${m.default ? ' [def:' + escapeHtml(m.default) + ']' : ''}${m.transform ? ' [' + escapeHtml(m.transform) + ']' : ''}</div>
      <button class="remove-btn" data-index="${i}" title="Remove">×</button>
      <button class="remove-btn" data-move-up="${i}" title="Move up" ${i === 0 ? 'disabled' : ''}>↑</button>
      <button class="remove-btn" data-move-down="${i}" title="Move down" ${i === state.pendingMappings.length - 1 ? 'disabled' : ''}>↓</button>
    </div>
  `).join('');
  list.querySelectorAll('.remove-btn[data-index]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.pendingMappings.splice(parseInt(btn.dataset.index), 1);
      renderMappingsList();
    });
  });
  list.querySelectorAll('[data-move-up]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const i = parseInt(btn.dataset.moveUp, 10);
      if (i > 0) {
        [state.pendingMappings[i - 1], state.pendingMappings[i]] = [state.pendingMappings[i], state.pendingMappings[i - 1]];
        renderMappingsList();
      }
    });
  });
  list.querySelectorAll('[data-move-down]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const i = parseInt(btn.dataset.moveDown, 10);
      if (i < state.pendingMappings.length - 1) {
        [state.pendingMappings[i], state.pendingMappings[i + 1]] = [state.pendingMappings[i + 1], state.pendingMappings[i]];
        renderMappingsList();
      }
    });
  });
}

// ── Fill Tab ────────────────────────────────────────────

let filterRecordsDebounce;
function setupFillTab() {
  $('recordSearch').addEventListener('input', () => {
    clearTimeout(filterRecordsDebounce);
    filterRecordsDebounce = setTimeout(filterRecords, 120);
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
  $('fillNextBtn').addEventListener('click', fillNextRecord);
  $('retryFailedBtn').addEventListener('click', () => triggerFill(false, true));
  $('undoFillBtn').addEventListener('click', undoLastFill);
  $('recordSelect').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); if (!$('fillBtn').disabled) triggerFill(false); return; }
    const sel = $('recordSelect');
    const opts = sel.options;
    if (opts.length === 0) return;
    if (e.key === 'ArrowDown') { e.preventDefault(); const i = Math.min(sel.selectedIndex + 1, opts.length - 1); sel.selectedIndex = i; state.selectedRecord = parseInt(opts[i].value, 10); updateFillButton(); updatePreviewAndWarnings(); saveToStorage(); }
    if (e.key === 'ArrowUp') { e.preventDefault(); const i = Math.max(sel.selectedIndex - 1, 0); sel.selectedIndex = i; state.selectedRecord = parseInt(opts[i].value, 10); updateFillButton(); updatePreviewAndWarnings(); saveToStorage(); }
  });
}

function updatePreviewAndWarnings() {
  const templateName = $('templateSelect').value;
  const previewWrap = $('previewBeforeFill');
  const previewList = $('previewList');
  const emptyWarn = $('emptyFieldsWarn');
  if (state.settings.showPreview === false && previewWrap) previewWrap.style.display = 'none';
  if (!templateName || !state.templates[templateName] || state.selectedRecord === null || !state.csvData) {
    if (previewWrap) previewWrap.style.display = 'none';
    if (emptyWarn) emptyWarn.style.display = 'none';
    return;
  }
  const record = state.csvData.rows[state.selectedRecord];
  const mappings = state.templates[templateName];
  const displayRecord = {};
  const emptyCols = [];
  mappings.forEach(m => {
    if (m.conditionColumn && m.conditionValue && String(record[m.conditionColumn] || '').trim() !== String(m.conditionValue).trim()) return;
    const val = getRecordValue(record, m.column, m);
    displayRecord[m.column] = val;
    if (state.settings.warnEmpty && !val) emptyCols.push(m.column);
  });
  if (state.settings.showPreview && previewList) {
    previewWrap.style.display = 'block';
    previewList.innerHTML = mappings.map(m => `<div>${escapeHtml(m.fieldLabel || m.fieldKey)} → ${escapeHtml(String(displayRecord[m.column] || '').slice(0, 40))}</div>`).join('');
  }
  if (emptyWarn) {
    if (emptyCols.length && state.settings.warnEmpty) {
      emptyWarn.textContent = 'Empty for this record: ' + emptyCols.join(', ');
      emptyWarn.style.display = 'block';
    } else emptyWarn.style.display = 'none';
  }
}

function undoLastFill() {
  if (!state.lastRestoreData || state.lastRestoreData.length === 0) { showToast('Nothing to undo', 'error'); return; }
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

function refreshFillHistory() {
  const wrap = $('fillHistoryWrap');
  const list = $('fillHistoryList');
  if (!wrap || !list) return;
  if (state.fillHistory.length === 0) { wrap.style.display = 'none'; return; }
  wrap.style.display = 'block';
  const fmt = (ts) => { const d = new Date(ts); return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); };
  list.innerHTML = state.fillHistory.slice(0, 5).map(h =>
    `<div>${h.templateName} · Record ${(h.recordIndex || 0) + 1} · ${fmt(h.timestamp)}</div>`
  ).join('');
}

function fillNextRecord() {
  if (!state.csvData || state.csvData.rows.length === 0) return;
  const next = state.selectedRecord === null ? 0 : Math.min(state.selectedRecord + 1, state.csvData.rows.length - 1);
  state.selectedRecord = next;
  state.recordPage = Math.floor(next / RECORD_PAGE_SIZE);
  refreshRecords();
  updateFillButton();
  updatePreviewAndWarnings();
  saveToStorage();
}

function filterRecords() {
  const query = $('recordSearch').value.toLowerCase();
  $('recordSelect').querySelectorAll('option').forEach((opt) => {
    opt.style.display = opt.textContent.toLowerCase().includes(query) ? '' : 'none';
  });
}

function triggerFill(fillAndNext, retryFailedOnly) {
  const templateName = $('templateSelect').value;
  if (!templateName || !state.templates[templateName]) { showToast('Select a template first', 'error'); return; }
  if (state.selectedRecord === null || !state.csvData) { showToast('Select a record first', 'error'); return; }

  const rawRecord = state.csvData.rows[state.selectedRecord];
  const mappings = state.templates[templateName];
  const displayRecord = {};
  mappings.forEach(m => {
    if (m.conditionColumn && m.conditionValue && String(rawRecord[m.conditionColumn] || '').trim() !== String(m.conditionValue).trim()) return;
    displayRecord[m.column] = getRecordValue(rawRecord, m.column, m);
  });

  const failedKeysOnly = retryFailedOnly && state.lastFillResult && state.lastFillResult.failedKeys && state.lastFillResult.failedKeys.length
    ? state.lastFillResult.failedKeys
    : null;

  $('fillBtn').disabled = true;
  $('fillAndNextBtn').disabled = true;
  $('fillBtn').innerHTML = '<div class="spinner"></div> Filling…';

  sendToActiveTab({ type: 'FILL_FIELDS', mappings, record: displayRecord, failedKeysOnly }, (response) => {
    $('fillBtn').disabled = false;
    $('fillBtn').textContent = '⚡ Autofill Document';
    updateFillButton();
    const errEl = $('fillErrors');
    const summaryEl = $('fillSummary');
    const retryBtn = $('retryFailedBtn');
    const staleHint = $('staleTabHint');
    const fillTabHint = $('fillTabHint');
    if (errEl) errEl.style.display = 'none';
    if (summaryEl) summaryEl.style.display = 'none';
    if (retryBtn) retryBtn.style.display = 'none';
    if (staleHint) staleHint.style.display = 'none';
    if (fillTabHint && state.isOnDocuSign) fillTabHint.style.display = 'none';

    if (chrome.runtime.lastError) {
      showToast('Reload the DocuSign tab and try again', 'error');
      if (staleHint) staleHint.style.display = 'block';
      if (fillTabHint) fillTabHint.style.display = 'block';
      return;
    }
    state.lastFillResult = response ? { count: response.count, errors: response.errors || [], failedKeys: response.failedKeys || [] } : null;
    state.lastRestoreData = response && response.previousValues && response.previousValues.length ? response.previousValues : null;
    const undoBtn = $('undoFillBtn');
    if (undoBtn) undoBtn.style.display = state.lastRestoreData ? 'block' : 'none';
    if (response && response.success && response.count > 0) {
      state.fillHistory.unshift({ recordIndex: state.selectedRecord, templateName, timestamp: Date.now() });
      if (state.fillHistory.length > FILL_HISTORY_MAX) state.fillHistory.pop();
      saveToStorage();
      refreshFillHistory();
    }

    if (response && response.success) {
      showToast(`Filled ${response.count} field(s)`, 'success');
      if (summaryEl) {
        summaryEl.style.display = 'block';
        summaryEl.innerHTML = '<strong>Filled ' + response.count + ' field(s)</strong>';
        if (response.errors && response.errors.length > 0) {
          summaryEl.innerHTML += '. <span style="color:var(--warning);">Failed: ' + escapeHtml(response.errors.slice(0, 3).join('; ')) + (response.errors.length > 3 ? ' (+' + (response.errors.length - 3) + ' more)' : '') + '</span>';
          if (retryBtn && response.failedKeys && response.failedKeys.length) retryBtn.style.display = 'block';
        }
        summaryEl.innerHTML += ' <a href="#" id="fillSummaryGotoMapping" style="color:var(--accent);">Re-capture in Mapping</a>';
        const link = document.getElementById('fillSummaryGotoMapping');
        if (link) link.addEventListener('click', (e) => { e.preventDefault(); document.querySelector('.tab[data-tab="mapping"]').click(); });
      }
      if (errEl && response.errors && response.errors.length > 0) {
        errEl.innerHTML = '<span style="color:var(--warning);">Some fields not found.</span> ' + escapeHtml(response.errors.slice(0, 2).join('; ')) + (response.errors.length > 2 ? ' (+' + (response.errors.length - 2) + ' more)' : '') + ' <a href="#" id="fillErrorsGotoMapping" style="color:var(--accent);">Re-capture in Mapping tab</a>';
        errEl.style.display = 'block';
        const link = document.getElementById('fillErrorsGotoMapping');
        if (link) link.addEventListener('click', (e) => { e.preventDefault(); document.querySelector('.tab[data-tab="mapping"]').click(); });
      }
      const doNext = fillAndNext || state.settings.batchFillMode;
      if (doNext) fillNextRecord();
    } else {
      const msg = response?.error || 'Fill failed — check console';
      showToast(msg, 'error');
      if (response && response.errors && response.errors.length > 0 && errEl) {
        errEl.innerHTML = '<span style="color:var(--danger);">' + escapeHtml(response.errors.slice(0, 3).join('; ')) + (response.errors.length > 3 ? ' (+' + (response.errors.length - 3) + ' more)' : '') + '</span>. <a href="#" id="fillErrorsGotoMapping2" style="color:var(--accent);">Re-capture in Mapping tab</a>';
        errEl.style.display = 'block';
        const link = document.getElementById('fillErrorsGotoMapping2');
        if (link) link.addEventListener('click', (e) => { e.preventDefault(); document.querySelector('.tab[data-tab="mapping"]').click(); });
      }
      if (response && response.failedKeys && response.failedKeys.length && retryBtn) retryBtn.style.display = 'block';
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
}

// ── Refresh / Render ────────────────────────────────────

function refreshAll() {
  refreshCSVStatus();
  refreshSheetPicker();
  refreshRecords();
  refreshFillTemplates();
  refreshSavedTemplates();
  refreshColumnDropdowns();
  updatePreviewAndWarnings();
  updateFillButton();
  refreshFillHistory();
}

function refreshSheetPicker() {
  const picker = $('sheetPicker');
  if (!picker) return;
  if (state.pendingXlsxSheets && state.pendingXlsxSheets.length > 0) {
    picker.style.display = 'block';
    const sel = $('sheetSelect');
    sel.innerHTML = state.pendingXlsxSheets.map((name, i) => `<option value="${i}">${escapeHtml(name)}</option>`).join('');
  } else {
    picker.style.display = 'none';
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
    $('columnsList').innerHTML = state.csvData.columns.map(col =>
      `<span style="padding:3px 7px;background:var(--surface2);border:1px solid var(--border);border-radius:4px;font-size:10px;color:var(--text-muted);">${escapeHtml(col)}</span>`
    ).join('');
    const cols = state.csvData.columns;
    const rows = state.csvData.rows.slice(0, 3);
    $('dataPreview').innerHTML = `<table style="border-collapse:collapse;width:100%;min-width:300px;">
      <thead><tr>${cols.map(c => `<th style="border:1px solid var(--border);padding:4px 6px;text-align:left;color:var(--text-dim);white-space:nowrap;">${escapeHtml(c)}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r => `<tr>${cols.map(c => `<td style="border:1px solid var(--border);padding:3px 6px;color:var(--text-muted);white-space:nowrap;max-width:80px;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(r[c] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;
  } else {
    dot.classList.remove('loaded');
    $('csvName').textContent = 'No file loaded';
    $('csvCount').textContent = '';
    $('fileInfo').style.display = 'none';
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
  const displayCol = cols.find(c => /name|provider|label|id|first/i.test(c)) || cols[0];
  const total = rows.length;
  if (total <= RECORD_PAGE_SIZE) {
    sel.innerHTML = rows.map((row, i) =>
      `<option value="${i}">${escapeHtml(String(row[displayCol] || '').trim() || `Row ${i + 1}`)}</option>`
    ).join('');
    if (paginationEl) paginationEl.style.display = 'none';
  } else {
    const start = state.recordPage * RECORD_PAGE_SIZE;
    const end = Math.min(start + RECORD_PAGE_SIZE, total);
    const pageRows = rows.slice(start, end);
    sel.innerHTML = pageRows.map((row, i) => {
      const idx = start + i;
      return `<option value="${idx}">${escapeHtml(String(row[displayCol] || '').trim() || `Row ${idx + 1}`)}</option>`;
    }).join('');
    if (paginationEl) {
      paginationEl.style.display = 'flex';
      paginationEl.innerHTML = `
        <span>${start + 1}–${end} of ${total}</span>
        <button type="button" class="btn btn-secondary btn-sm" id="recordPrevPage" ${state.recordPage === 0 ? 'disabled' : ''}>Prev</button>
        <button type="button" class="btn btn-secondary btn-sm" id="recordNextPage" ${end >= total ? 'disabled' : ''}>Next</button>
      `;
      const prevBtn = $('recordPrevPage');
      const nextBtn = $('recordNextPage');
      if (prevBtn) prevBtn.addEventListener('click', () => { state.recordPage = Math.max(0, state.recordPage - 1); refreshRecords(); updateFillButton(); });
      if (nextBtn) nextBtn.addEventListener('click', () => { state.recordPage = Math.min(Math.floor((total - 1) / RECORD_PAGE_SIZE), state.recordPage + 1); refreshRecords(); updateFillButton(); });
    }
  }
  if (progressEl) progressEl.textContent = total ? `(${total} records)` : '';
  if (state.selectedRecord !== null && state.selectedRecord >= 0 && state.selectedRecord < total) {
    if (total > RECORD_PAGE_SIZE) {
      const start = state.recordPage * RECORD_PAGE_SIZE;
      const end = Math.min(start + RECORD_PAGE_SIZE, total);
      if (state.selectedRecord >= start && state.selectedRecord < end) sel.value = String(state.selectedRecord);
      progressEl.textContent = `(Record ${state.selectedRecord + 1} of ${total})`;
    } else {
      sel.value = String(state.selectedRecord);
      progressEl.textContent = `(Record ${state.selectedRecord + 1} of ${total})`;
    }
  }
}

function refreshFillTemplates() {
  const sel = $('templateSelect');
  const names = Object.keys(state.templates);
  sel.innerHTML = '<option value="">— Select a template —</option>' +
    names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)} (${state.templates[n].length} fields)</option>`).join('');
}

function refreshSavedTemplates() {
  const list = $('savedTemplatesList');
  const names = Object.keys(state.templates);
  $('templateBadge').textContent = `${names.length} saved`;
  if (names.length === 0) {
    list.innerHTML = `<div class="empty-state" style="padding:16px;"><div class="empty-text">No templates saved yet</div></div>`;
    return;
  }
  list.innerHTML = names.map(name => `
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
  `).join('');
  list.querySelectorAll('[data-load]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.pendingMappings = [...state.templates[btn.dataset.load]];
      $('templateNameInput').value = btn.dataset.load;
      renderMappingsList();
      document.querySelector('.tab[data-tab="mapping"]').click();
    });
  });
  list.querySelectorAll('[data-clone]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const base = btn.dataset.clone;
      let newName = 'Copy of ' + base;
      let c = 1;
      while (state.templates[newName]) { newName = 'Copy of ' + base + ' (' + (++c) + ')'; }
      state.templates[newName] = [...state.templates[base]];
      saveToStorage();
      refreshSavedTemplates();
      refreshFillTemplates();
      showToast('Cloned as "' + newName + '"', 'success');
    });
  });
  list.querySelectorAll('[data-rename]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const oldName = btn.dataset.rename;
      const newName = window.prompt('New template name:', oldName);
      if (!newName || newName.trim() === '') return;
      const trimmed = newName.trim();
      if (trimmed === oldName) return;
      if (state.templates[trimmed]) { showToast('A template with that name already exists', 'error'); return; }
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

function refreshColumnDropdowns() {
  const sel = $('columnSelect');
  if (!state.csvData) { sel.innerHTML = '<option value="">— Load CSV first —</option>'; return; }
  sel.innerHTML = '<option value="">— Select column —</option>' +
    state.csvData.columns.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
}

// ── DocuSign Detection ──────────────────────────────────

function checkDocuSignTab() {
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
        });
      } catch (e) {
        state.contentScriptReady = false;
        $('staleTabHint').style.display = 'block';
      }
    } else {
      state.isOnDocuSign = false;
      state.contentScriptReady = false;
      $('pageStatus').textContent = 'Not on DocuSign';
      $('pageStatus').style.color = 'var(--text-dim)';
      $('fillTabHint').style.display = 'block';
      $('staleTabHint').style.display = 'none';
    }
    updateFillButton();
  });
}

// ── Messaging ───────────────────────────────────────────

function sendToActiveTab(message, callback) {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (!tabs[0]) { if (callback) callback({ error: 'No active tab' }); return; }
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
      document.querySelector('.tab[data-tab="mapping"]').click();
    }
  });
}

// ── Toast ───────────────────────────────────────────────

let toastTimeout;
function showToast(msg, type = '') {
  const toast = $('toast');
  toast.textContent = msg;
  toast.className = 'toast show' + (type ? ' ' + type : '');
  clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => toast.classList.remove('show'), 2500);
}
