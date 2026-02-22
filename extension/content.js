// ═══════════════════════════════════════════════════════
//  DocuFill – content.js
//  Injected into DocuSign pages (all frames)
// ═══════════════════════════════════════════════════════

let captureMode = false;
let highlightedEl = null;
let overlay = null;

function createOverlay() {
  if (overlay) return;
  function append() {
    if (!document.body) return;
    overlay = document.createElement('div');
    overlay.id = '__docufill_overlay';
    overlay.style.cssText = `
      position: fixed; top: 12px; left: 50%; transform: translateX(-50%);
      background: rgba(79,124,255,0.92); color: #fff;
      padding: 8px 18px; border-radius: 20px;
      font-family: -apple-system, sans-serif; font-size: 13px; font-weight: 600;
      z-index: 999999; pointer-events: none;
      box-shadow: 0 4px 20px rgba(79,124,255,0.4);
    `;
    overlay.textContent = '🎯 DocuFill: Click any field to capture it';
    document.body.appendChild(overlay);
  }
  if (document.body) append();
  else document.addEventListener('DOMContentLoaded', append);
}

function removeOverlay() {
  if (overlay) { overlay.remove(); overlay = null; }
}

// Iframe fill: only top frame orchestrates; child frames respond to postMessage
const DOCUFILL_PREFIX = '__docufill_';
function isTopFrame() { try { return window.self === window.top; } catch (e) { return false; } }

function getFieldValue(el) {
  const tag = el.tagName.toLowerCase();
  const type = (el.type || '').toLowerCase();
  if (tag === 'select') return el.options[el.selectedIndex]?.value ?? '';
  if (type === 'checkbox' || type === 'radio') return el.checked ? 'true' : 'false';
  return el.value ?? '';
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'PING') {
    sendResponse({ ok: true });
    return true;
  }
  if (message.type === 'SCAN_FIELDS') {
    if (isTopFrame()) {
      const allFields = getFieldsInDocument(document);
      const iframes = document.querySelectorAll('iframe');
      if (iframes.length === 0) {
        sendResponse({ success: true, fields: allFields });
        return true;
      }
      const results = [allFields];
      let pending = iframes.length;
      function onScanResult(ev) {
        if (ev.data && ev.data.type === DOCUFILL_PREFIX + 'SCAN_RESULT') {
          results.push(ev.data.fields || []);
          pending--;
          if (pending <= 0) finish();
        }
      }
      function finish() {
        window.removeEventListener('message', onScanResult);
        const merged = [];
        const seen = {};
        results.forEach(function (arr) {
          (arr || []).forEach(function (f) {
            if (f.fieldKey && !seen[f.fieldKey]) { seen[f.fieldKey] = true; merged.push(f); }
          });
        });
        sendResponse({ success: true, fields: merged });
      }
      window.addEventListener('message', onScanResult);
      iframes.forEach(function (frame) {
        try { frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'SCAN_FIELDS' }, '*'); } catch (e) { pending--; }
      });
      setTimeout(function () { if (pending > 0) { pending = 0; finish(); } }, 3000);
    } else {
      sendResponse({ success: true, fields: getFieldsInDocument(document) });
    }
    return true;
  }
  if (message.type === 'RESTORE_FIELDS') {
    const restores = message.restores || [];
    let count = 0;
    restores.forEach(function (r) {
      const el = findFieldByKey(r.fieldKey);
      if (el) { try { setFieldValue(el, r.value); count++; } catch (e) {} }
    });
    const iframes = document.querySelectorAll('iframe');
    iframes.forEach(function (frame) {
      try { frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'RESTORE', restores: restores }, '*'); } catch (e) {}
    });
    sendResponse({ success: true, count });
    return true;
  }
  if (message.type === 'CAPTURE_MODE_ON') activateCaptureMode();
  if (message.type === 'CAPTURE_MODE_OFF') deactivateCaptureMode();
  if (message.type === 'FILL_FIELDS') {
    const mappings = message.mappings;
    const record = message.record;
    const failedKeysOnly = message.failedKeysOnly || null;
    if (isTopFrame()) {
      const mainResult = fillFields(mappings, record, failedKeysOnly);
      const iframes = document.querySelectorAll('iframe');
      if (iframes.length === 0) {
        sendResponse(mainResult);
        return true;
      }
      const results = [mainResult];
      let pending = iframes.length;
      function onIframeResult(ev) {
        if (ev.data && ev.data.type === DOCUFILL_PREFIX + 'FILL_RESULT') {
          results.push(ev.data.result);
          pending--;
          if (pending <= 0) finish();
        }
      }
      function finish() {
        window.removeEventListener('message', onIframeResult);
        const merged = { success: false, count: 0, errors: [], failedKeys: [], previousValues: [], error: null };
        results.forEach(function (r) {
          merged.count += r.count || 0;
          if (r.errors) merged.errors = merged.errors.concat(r.errors);
          if (r.failedKeys) merged.failedKeys = merged.failedKeys.concat(r.failedKeys);
          if (r.previousValues) merged.previousValues = merged.previousValues.concat(r.previousValues);
        });
        merged.success = merged.errors.length === 0 || merged.count > 0;
        merged.error = merged.errors[0] || null;
        sendResponse(merged);
      }
      window.addEventListener('message', onIframeResult);
      iframes.forEach(function (frame) {
        try {
          frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'FILL', mappings: mappings, record: record, failedKeysOnly: failedKeysOnly }, '*');
        } catch (e) { pending--; }
      });
      setTimeout(function () {
        if (pending > 0) { pending = 0; finish(); }
      }, 500);
    } else {
      sendResponse(fillFields(mappings, record, failedKeysOnly));
    }
  }
  return true;
});

window.addEventListener('message', function (ev) {
  if (!ev.data) return;
  if (ev.data.type === DOCUFILL_PREFIX + 'FILL') {
    const result = fillFields(ev.data.mappings || [], ev.data.record || {}, ev.data.failedKeysOnly || null);
    try { ev.source.postMessage({ type: DOCUFILL_PREFIX + 'FILL_RESULT', result: result }, '*'); } catch (e) {}
  }
  if (ev.data.type === DOCUFILL_PREFIX + 'RESTORE') {
    const restores = ev.data.restores || [];
    restores.forEach(function (r) {
      const el = findFieldByKey(r.fieldKey);
      if (el) { try { setFieldValue(el, r.value); } catch (e) {} }
    });
  }
  if (ev.data.type === DOCUFILL_PREFIX + 'SCAN_FIELDS') {
    const fields = getFieldsInDocument(document);
    try { ev.source.postMessage({ type: DOCUFILL_PREFIX + 'SCAN_RESULT', fields: fields }, '*'); } catch (e) {}
  }
});

function getFieldsInDocument(doc) {
  const seen = {};
  const out = [];
  const root = doc.body || doc.documentElement;
  if (!root) return out;
  const candidates = root.querySelectorAll('input, select, textarea, [contenteditable="true"], [role="textbox"], .ds-field, [data-testid*="field"]');
  candidates.forEach(function (el) {
    const inputEl = el.matches('input, select, textarea, [contenteditable="true"], [role="textbox"]') ? el : (el.querySelector('input, select, textarea') || (findInputElement(el) || el));
    if (!inputEl || inputEl.tagName === 'INPUT' && (inputEl.type === 'hidden' || inputEl.type === 'submit' || inputEl.type === 'button' || inputEl.type === 'image')) return;
    const info = extractFieldInfo(inputEl);
    const key = info.key;
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push({ fieldKey: key, fieldLabel: (info.label || '').trim() || key });
  });
  return out;
}

function activateCaptureMode() {
  captureMode = true;
  createOverlay();
  document.addEventListener('mouseover', onMouseOver);
  document.addEventListener('mouseout', onMouseOut);
  document.addEventListener('click', onClick, true);
}

function deactivateCaptureMode() {
  captureMode = false;
  removeOverlay();
  clearHighlight();
  document.removeEventListener('mouseover', onMouseOver);
  document.removeEventListener('mouseout', onMouseOut);
  document.removeEventListener('click', onClick, true);
}

function onMouseOver(e) {
  if (!captureMode) return;
  const el = findInputElement(e.target);
  if (!el) return;
  clearHighlight();
  highlightedEl = el;
  el.__docufill_origOutline = el.style.outline;
  el.style.outline = '2px solid #4f7cff';
  el.style.outlineOffset = '2px';
}

function onMouseOut() {
  if (!captureMode) return;
  clearHighlight();
}

function clearHighlight() {
  if (highlightedEl) {
    highlightedEl.style.outline = highlightedEl.__docufill_origOutline || '';
    highlightedEl.style.outlineOffset = '';
    highlightedEl = null;
  }
}

function onClick(e) {
  if (!captureMode) return;
  const el = findInputElement(e.target);
  if (!el) return;
  e.preventDefault();
  e.stopPropagation();
  const fieldInfo = extractFieldInfo(el);
  chrome.runtime.sendMessage({ type: 'FIELD_CAPTURED', fieldKey: fieldInfo.key, fieldLabel: fieldInfo.label });
  clearHighlight();
  el.style.outline = '2px solid #22c55e';
  setTimeout(() => { el.style.outline = el.__docufill_origOutline || ''; }, 800);
}

function fillFields(mappings, record, failedKeysOnly) {
  let count = 0;
  const errors = [];
  const failedKeys = [];
  const previousValues = [];
  const toProcess = failedKeysOnly && failedKeysOnly.length
    ? mappings.filter(m => failedKeysOnly.indexOf(m.fieldKey) >= 0)
    : mappings;
  toProcess.forEach((mapping) => {
    let value = record[mapping.column];
    if (value === undefined || value === null) return;
    const el = findFieldByKey(mapping.fieldKey);
    if (!el) { errors.push('Field not found: ' + (mapping.fieldLabel || mapping.fieldKey)); failedKeys.push(mapping.fieldKey); return; }
    try {
      previousValues.push({ fieldKey: mapping.fieldKey, value: getFieldValue(el) });
      setFieldValue(el, String(value));
      count++;
    } catch (err) { errors.push('Failed to fill ' + (mapping.fieldLabel || mapping.fieldKey) + ': ' + err.message); failedKeys.push(mapping.fieldKey); }
  });
  return { success: errors.length === 0 || count > 0, count, errors, failedKeys, previousValues, error: errors[0] || null };
}

function setFieldValue(el, value) {
  const tag = el.tagName.toLowerCase();
  const type = (el.type || '').toLowerCase();
  if (tag === 'select') {
    const option = Array.from(el.options).find(
      o => o.value === value || o.text === value ||
           o.value.toLowerCase() === value.toLowerCase() ||
           o.text.toLowerCase() === value.toLowerCase()
    );
    if (option) { el.value = option.value; triggerEvents(el, ['change', 'input', 'blur']); }
  } else if (type === 'checkbox') {
    el.checked = /true|yes|1|x/i.test(value);
    triggerEvents(el, ['change', 'click']);
  } else if (type === 'radio') {
    if (el.value.toLowerCase() === value.toLowerCase()) {
      el.checked = true;
      triggerEvents(el, ['change', 'click']);
    }
  } else {
    const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
    const nativeTextareaSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;
    const setter = tag === 'textarea' ? nativeTextareaSetter : nativeInputValueSetter;
    if (setter) setter.call(el, value);
    else el.value = value;
    triggerEvents(el, ['input', 'change', 'blur', 'keyup']);
  }
}

function triggerEvents(el, events) {
  events.forEach(eventName => el.dispatchEvent(new Event(eventName, { bubbles: true, cancelable: true })));
}

function findInputElement(el) {
  let target = el;
  let depth = 0;
  while (target && depth < 5) {
    const tag = target.tagName?.toLowerCase();
    if (['input', 'select', 'textarea'].includes(tag)) return target;
    if (target.contentEditable === 'true') return target;
    if (target.classList && (
      target.classList.contains('ds-field') ||
      target.getAttribute('data-testid')?.includes('field') ||
      target.getAttribute('role') === 'textbox'
    )) return target;
    target = target.parentElement;
    depth++;
  }
  return null;
}

function extractFieldInfo(el) {
  const id = el.id;
  const name = el.name || el.getAttribute('name');
  const dataId = el.getAttribute('data-id') || el.getAttribute('data-field-id') ||
                 el.getAttribute('data-tab-id') || el.getAttribute('data-testid');
  const ariaLabel = el.getAttribute('aria-label');
  let label = '';
  if (id) { const labelEl = document.querySelector(`label[for="${id}"]`); if (labelEl) label = labelEl.textContent.trim(); }
  if (!label && ariaLabel) label = ariaLabel;
  if (!label && el.placeholder) label = el.placeholder;
  if (!label && name) label = name;
  if (!label) label = el.className.split(' ')[0] || 'Unknown Field';
  const key = id || name || dataId || buildXPath(el);
  return { key, label };
}

function findFieldByKey(key) {
  let el = document.getElementById(key);
  if (el) return el;
  el = document.querySelector(`[name="${key}"]`);
  if (el) return el;
  el = document.querySelector(`[data-id="${key}"]`) ||
       document.querySelector(`[data-field-id="${key}"]`) ||
       document.querySelector(`[data-tab-id="${key}"]`) ||
       document.querySelector(`[data-testid="${key}"]`);
  if (el) return el;
  try {
    const result = document.evaluate(key, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
    if (result.singleNodeValue) return result.singleNodeValue;
  } catch (e) {}
  return null;
}

function buildXPath(el) {
  if (el.id) return `//*[@id="${el.id}"]`;
  const parts = [];
  let current = el;
  while (current && current.nodeType === Node.ELEMENT_NODE) {
    let idx = 1;
    let sibling = current.previousSibling;
    while (sibling) { if (sibling.nodeType === Node.ELEMENT_NODE && sibling.tagName === current.tagName) idx++; sibling = sibling.previousSibling; }
    parts.unshift(`${current.tagName.toLowerCase()}[${idx}]`);
    current = current.parentElement;
    if (parts.length > 6) break;
  }
  return '/' + parts.join('/');
}
