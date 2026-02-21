// ═══════════════════════════════════════════════════════
//  DocuFill – popup.js
// ═══════════════════════════════════════════════════════

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
};

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
  setupTabs();
  setupDataTab();
  setupMappingTab();
  setupFillTab();
  checkDocuSignTab();
  listenForContentMessages();
});

// ── Storage ─────────────────────────────────────────────

async function loadFromStorage() {
  return new Promise((resolve) => {
    chrome.storage.local.get(['csvData', 'templates', 'lastTemplate', 'lastRecord'], (result) => {
      if (result.csvData) state.csvData = result.csvData;
      if (result.templates) state.templates = result.templates;
      refreshAll();
      if (result.lastTemplate && state.templates[result.lastTemplate]) {
        $('templateSelect').value = result.lastTemplate;
        state.selectedRecord = null;
      }
      if (result.lastRecord != null && state.csvData && result.lastRecord >= 0 && result.lastRecord < state.csvData.rows.length) {
        state.selectedRecord = result.lastRecord;
        $('recordSelect').value = String(result.lastRecord);
      }
      updateFillButton();
      resolve();
    });
  });
}

function saveToStorage() {
  const payload = { csvData: state.csvData, templates: state.templates };
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
      document.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
      document.querySelectorAll('.panel').forEach((p) => p.classList.remove('active'));
      tab.classList.add('active');
      $('tab-' + tab.dataset.tab).classList.add('active');
    });
  });
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
  $('captureModeBtn').addEventListener('click', toggleCaptureMode);
  $('confirmMappingBtn').addEventListener('click', confirmMapping);
  $('cancelCaptureBtn').addEventListener('click', cancelCapture);
  $('saveTemplateBtn').addEventListener('click', saveTemplate);
  $('newTemplateBtn').addEventListener('click', newTemplate);
  $('loadTemplateBtn').addEventListener('click', loadTemplateIntoEditor);
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
  state.pendingMappings.push({
    fieldKey: state.pendingField.fieldKey,
    fieldLabel: state.pendingField.fieldLabel,
    column: colSelect.value,
  });
  state.pendingField = null;
  $('pendingCapture').style.display = 'none';
  renderMappingsList();
  showToast('Mapping added ✓', 'success');
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
    <div class="mapping-row">
      <div class="field-tag" title="${escapeHtml(m.fieldKey)}">${escapeHtml(m.fieldLabel || m.fieldKey)}</div>
      <div class="mapping-arrow">→</div>
      <div class="field-tag" style="color:var(--accent);">${escapeHtml(m.column)}</div>
      <button class="remove-btn" data-index="${i}">×</button>
    </div>
  `).join('');
  list.querySelectorAll('.remove-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.pendingMappings.splice(parseInt(btn.dataset.index), 1);
      renderMappingsList();
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
    state.selectedRecord = e.target.value ? parseInt(e.target.value) : null;
    updateFillButton();
    saveToStorage();
  });
  $('templateSelect').addEventListener('change', () => {
    updateFillButton();
    saveToStorage();
  });
  $('fillBtn').addEventListener('click', triggerFill);
  $('fillNextBtn').addEventListener('click', fillNextRecord);
}
function fillNextRecord() {
  if (!state.csvData || state.csvData.rows.length === 0) return;
  const next = state.selectedRecord === null ? 0 : Math.min(state.selectedRecord + 1, state.csvData.rows.length - 1);
  state.selectedRecord = next;
  $('recordSelect').value = String(next);
  updateFillButton();
  saveToStorage();
  if (state.isOnDocuSign && $('templateSelect').value) triggerFill();
}

function filterRecords() {
  const query = $('recordSearch').value.toLowerCase();
  $('recordSelect').querySelectorAll('option').forEach((opt) => {
    opt.style.display = opt.textContent.toLowerCase().includes(query) ? '' : 'none';
  });
}

function triggerFill() {
  const templateName = $('templateSelect').value;
  if (!templateName || !state.templates[templateName]) { showToast('Select a template first', 'error'); return; }
  if (state.selectedRecord === null || !state.csvData) { showToast('Select a record first', 'error'); return; }

  const record = state.csvData.rows[state.selectedRecord];
  const mappings = state.templates[templateName];

  $('fillBtn').disabled = true;
  $('fillBtn').innerHTML = '<div class="spinner"></div> Filling…';

  sendToActiveTab({ type: 'FILL_FIELDS', mappings, record }, (response) => {
    $('fillBtn').disabled = false;
    $('fillBtn').textContent = '⚡ Autofill Document';
    const errEl = $('fillErrors');
    if (errEl) errEl.style.display = 'none';
    if (chrome.runtime.lastError) {
      showToast('Open a DocuSign document in this tab first', 'error');
      return;
    }
    if (response && response.success) {
      showToast(`Filled ${response.count} field(s)`, 'success');
      if (response.errors && response.errors.length > 0) {
        if (errEl) {
          errEl.innerHTML = '<span style="color:var(--warning);">Some fields not found.</span> ' +
            escapeHtml(response.errors.slice(0, 2).join('; ')) +
            (response.errors.length > 2 ? ' (+' + (response.errors.length - 2) + ' more)' : '') +
            ' <a href="#" id="fillErrorsGotoMapping" style="color:var(--accent);">Re-capture in Mapping tab</a>';
          errEl.style.display = 'block';
          const link = document.getElementById('fillErrorsGotoMapping');
          if (link) link.addEventListener('click', (e) => { e.preventDefault(); document.querySelector('.tab[data-tab="mapping"]').click(); });
        }
      }
    } else {
      const msg = response?.error || 'Fill failed — check console';
      showToast(msg, 'error');
      if (response && response.errors && response.errors.length > 0 && errEl) {
        errEl.innerHTML = '<span style="color:var(--danger);">' + escapeHtml(response.errors.slice(0, 3).join('; ')) + (response.errors.length > 3 ? ' (+' + (response.errors.length - 3) + ' more)' : '') + '</span>. <a href="#" id="fillErrorsGotoMapping2" style="color:var(--accent);">Re-capture in Mapping tab</a>';
        errEl.style.display = 'block';
        const link = document.getElementById('fillErrorsGotoMapping2');
        if (link) link.addEventListener('click', (e) => { e.preventDefault(); document.querySelector('.tab[data-tab="mapping"]').click(); });
      }
    }
  });
}

function updateFillButton() {
  const hasTemplate = !!$('templateSelect').value;
  const hasRecord = state.selectedRecord !== null;
  $('fillBtn').disabled = !(hasTemplate && hasRecord && state.isOnDocuSign);
}

// ── Refresh / Render ────────────────────────────────────

function refreshAll() {
  refreshCSVStatus();
  refreshSheetPicker();
  refreshRecords();
  refreshFillTemplates();
  refreshSavedTemplates();
  refreshColumnDropdowns();
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
  if (!state.csvData) { sel.innerHTML = '<option>Load a CSV first</option>'; return; }
  const cols = state.csvData.columns;
  const displayCol = cols.find(c => /name|provider|label|id|first/i.test(c)) || cols[0];
  sel.innerHTML = state.csvData.rows.map((row, i) =>
    `<option value="${i}">${escapeHtml(row[displayCol] || `Row ${i + 1}`)}</option>`
  ).join('');
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
        <button class="btn btn-secondary btn-sm" data-load="${escapeHtml(name)}">Edit</button>
        <button class="btn btn-danger btn-sm" data-delete="${escapeHtml(name)}">✕</button>
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
  list.querySelectorAll('[data-delete]').forEach((btn) => {
    btn.addEventListener('click', () => {
      delete state.templates[btn.dataset.delete];
      saveToStorage();
      refreshSavedTemplates();
      refreshFillTemplates();
      showToast(`Deleted "${btn.dataset.delete}"`);
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
    } else {
      state.isOnDocuSign = false;
      $('pageStatus').textContent = 'Not on DocuSign';
      $('pageStatus').style.color = 'var(--text-dim)';
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
