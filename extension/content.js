// ═══════════════════════════════════════════════════════
//  DocuFill – content.js
//  Injected into DocuSign pages (all frames)
// ═══════════════════════════════════════════════════════

// Immediate, highly visible log - try multiple methods
try {
  console.log('%c🔵🔵🔵 DOCUFILL CONTENT SCRIPT LOADED 🔵🔵🔵', 'color: #4f7cff; font-size: 18px; font-weight: bold; background: #1a1d27; padding: 8px; border: 2px solid #4f7cff; border-radius: 4px;');
  console.log('DocuFill: Frame URL:', window.location.href);
  console.log('DocuFill: Is top frame:', window.self === window.top);
  console.log('DocuFill: If you see this message, the content script is working!');
  console.warn('⚠️ DocuFill: This is a WARNING level message - should be visible');
  console.error('❌ DocuFill: This is an ERROR level message - should definitely be visible');
} catch (e) {
  // Fallback if console is blocked
  window.__docufill_loaded = true;
}

// Set a global flag so we can test if script loaded
window.__docufill_script_loaded = true;
window.__docufill_frame_url = window.location.href;
window.__docufill_is_top_frame = window.self === window.top;

let captureMode = false;
let highlightedEl = null;
let overlay = null;

function createOverlay() {
  if (overlay) return;
  console.log('DocuFill: Creating overlay banner...');
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
    console.log('DocuFill: Overlay banner created and displayed');
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
  if (el.isContentEditable || el.getAttribute('contenteditable') === 'true') return (el.textContent || '').trim();
  return el.value ?? '';
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'PING') {
    sendResponse({ ok: true });
    return true;
  }
  if (message.type === 'EXTRACT_DOCUSIGN_FIELDS') {
    const fields = extractDocuSignInternalFields();
    sendResponse({ success: true, fields: fields });
    return true;
  }
  if (message.type === 'READ_PROPERTY_PANEL') {
    const fields = readPropertyPanelFields();
    sendResponse({ success: true, fields: fields });
    return true;
  }
  if (message.type === 'READ_PROPERTY_PANEL') {
    const fields = readPropertyPanelFields();
    sendResponse({ success: true, fields: fields });
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
      let retries = 0;
      const maxRetries = 2;
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
      function retryIframes() {
        if (retries < maxRetries && pending > 0) {
          retries++;
          iframes.forEach(function (frame) {
            try { 
              frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'SCAN_FIELDS' }, '*'); 
            } catch (e) { 
              pending--; 
            }
          });
          setTimeout(function () { 
            if (pending > 0) retryIframes();
            else finish();
          }, 2000);
        } else {
          finish();
        }
      }
      window.addEventListener('message', onScanResult);
      iframes.forEach(function (frame) {
        try { 
          frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'SCAN_FIELDS' }, '*'); 
        } catch (e) { 
          pending--; 
        }
      });
      // Wait longer for iframes to respond (especially for PDF viewers)
      setTimeout(function () { 
        if (pending > 0) retryIframes();
        else finish();
      }, 5000);
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
  if (message.type === 'CAPTURE_MODE_ON') {
    activateCaptureMode();
    if (isTopFrame()) broadcastToFrames(DOCUFILL_PREFIX + 'CAPTURE_MODE_ON');
  }
  if (message.type === 'CAPTURE_MODE_OFF') {
    deactivateCaptureMode();
    if (isTopFrame()) broadcastToFrames(DOCUFILL_PREFIX + 'CAPTURE_MODE_OFF');
  }
  if (message.type === 'VERIFY_MAPPINGS') {
    const fieldKeys = message.fieldKeys || [];
    function runVerify() {
      const found = [];
      const missing = [];
      fieldKeys.forEach(function (key) {
        const el = findFieldByKey(key);
        if (el) found.push(key); else missing.push(key);
      });
      return { found, missing };
    }
    if (!isTopFrame()) {
      sendResponse({ success: true, found: runVerify().found, missing: runVerify().missing });
      return true;
    }
    const main = runVerify();
    const iframes = document.querySelectorAll('iframe');
    if (iframes.length === 0) {
      sendResponse({ success: true, found: main.found, missing: main.missing });
      return true;
    }
    const allFound = {};
    main.found.forEach(function (k) { allFound[k] = true; });
    let pending = iframes.length;
    function onResult(ev) {
      if (ev.data && ev.data.type === DOCUFILL_PREFIX + 'VERIFY_RESULT') {
        (ev.data.found || []).forEach(function (k) { allFound[k] = true; });
        pending--;
        if (pending <= 0) finish();
      }
    }
    function finish() {
      window.removeEventListener('message', onResult);
      const found = Object.keys(allFound);
      const missing = fieldKeys.filter(function (k) { return !allFound[k]; });
      sendResponse({ success: true, found, missing });
    }
    window.addEventListener('message', onResult);
    iframes.forEach(function (frame) {
      try { frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'VERIFY_MAPPINGS', fieldKeys: fieldKeys }, '*'); } catch (e) { pending--; }
    });
    setTimeout(function () { if (pending > 0) { pending = 0; finish(); } }, 2000);
    return true;
  }
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

function broadcastToFrames(type) {
  try {
    const iframes = document.querySelectorAll('iframe');
    iframes.forEach(function (frame) {
      try { frame.contentWindow.postMessage({ type: type }, '*'); } catch (e) {}
    });
  } catch (e) {}
}

window.addEventListener('message', function (ev) {
  if (!ev.data) return;
  if (ev.data.type === DOCUFILL_PREFIX + 'CAPTURE_MODE_ON') { 
    activateCaptureMode(); 
    // Also broadcast to any nested iframes
    if (isTopFrame()) {
      setTimeout(() => {
        const iframes = document.querySelectorAll('iframe');
        iframes.forEach(function (frame) {
          try { 
            frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'CAPTURE_MODE_ON' }, '*'); 
          } catch (e) {}
        });
      }, 100);
    }
    return; 
  }
  if (ev.data.type === DOCUFILL_PREFIX + 'CAPTURE_MODE_OFF') { 
    deactivateCaptureMode(); 
    // Also broadcast to any nested iframes
    if (isTopFrame()) {
      const iframes = document.querySelectorAll('iframe');
      iframes.forEach(function (frame) {
        try { 
          frame.contentWindow.postMessage({ type: DOCUFILL_PREFIX + 'CAPTURE_MODE_OFF' }, '*'); 
        } catch (e) {}
      });
    }
    return; 
  }
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
  if (ev.data.type === DOCUFILL_PREFIX + 'VERIFY_MAPPINGS') {
    const fieldKeys = ev.data.fieldKeys || [];
    const found = [];
    const missing = [];
    fieldKeys.forEach(function (key) {
      const el = findFieldByKey(key);
      if (el) found.push(key); else missing.push(key);
    });
    try { ev.source.postMessage({ type: DOCUFILL_PREFIX + 'VERIFY_RESULT', found: found, missing: missing }, '*'); } catch (e) {}
  }
});

function getFieldsInDocument(doc) {
  const seen = {};
  const out = [];
  const root = doc.body || doc.documentElement;
  if (!root) return out;
  
  // Strategy 1: Standard form elements
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
  
  // Strategy 2: DocuSign prepare mode - elements with data-tab-id, data-tab-label, or data-field-id
  const docusignFields = root.querySelectorAll('[data-tab-id], [data-tab-label], [data-field-id]');
  docusignFields.forEach(function (el) {
    const tabId = el.getAttribute('data-tab-id') || el.getAttribute('data-field-id');
    if (!tabId || seen[tabId]) return;
    // Check if this is actually a field (not just any element with these attributes)
    const hasFieldIndicator = el.getAttribute('data-tab-label') || 
                              el.getAttribute('data-field-id') ||
                              el.classList.contains('ds-field') ||
                              el.querySelector('input, select, textarea, [contenteditable="true"]') ||
                              el.getAttribute('role') === 'textbox' ||
                              el.getAttribute('tabindex') !== null;
    if (!hasFieldIndicator) return;
    const info = extractFieldInfo(el);
    const key = info.key || tabId;
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push({ fieldKey: key, fieldLabel: (info.label || '').trim() || key });
  });
  
  // Strategy 3: Shadow DOM traversal
  try {
    const allElements = root.querySelectorAll('*');
    allElements.forEach(function (el) {
      if (el.shadowRoot) {
        const shadowFields = getFieldsInDocument(el.shadowRoot);
        shadowFields.forEach(function (f) {
          if (f.fieldKey && !seen[f.fieldKey]) {
            seen[f.fieldKey] = true;
            out.push(f);
          }
        });
      }
    });
  } catch (e) {}
  
  // Strategy 4: Try to access DocuSign's internal field data (React state, window objects, Redux store)
  try {
    // Check for various DocuSign global objects
    const possibleGlobals = [
      window.__DOCUSIGN_FIELDS__,
      window.DocuSign?.fields,
      window.DocuSign?.tabs,
      window.__DS_FIELDS__,
      window.dsFields,
      window.__REDUX_STORE__,
      window.store,
      window.__REACT_DEVTOOLS_GLOBAL_HOOK__?.renderers?.get(1)?.currentDispatcherRef,
      window.__REACT_DEVTOOLS_GLOBAL_HOOK__?.renderers?.get(1)?.findFiberByHostInstance
    ];
    
    for (let i = 0; i < possibleGlobals.length; i++) {
      const data = possibleGlobals[i];
      if (!data) continue;
      
      // If it's a Redux store, try to get state
      if (data && typeof data.getState === 'function') {
        try {
          const state = data.getState();
          // Look for tabs/fields in common Redux state locations
          const tabs = state?.tabs || state?.fields || state?.document?.tabs || state?.envelope?.tabs || 
                       state?.prepare?.tabs || state?.signing?.tabs;
          if (tabs && Array.isArray(tabs)) {
            tabs.forEach(function (f) {
              const key = f.tabId || f.id || f.fieldId || f.key || f.recipientId + '_' + f.documentId + '_' + f.pageNumber + '_' + f.xPosition + '_' + f.yPosition;
              const label = f.tabLabel || f.label || f.name || f.type || '';
              if (key && !seen[key]) {
                seen[key] = true;
                out.push({ fieldKey: String(key), fieldLabel: label || String(key) });
              }
            });
          }
        } catch (e) {}
      }
      
      // If it's an array of fields
      if (Array.isArray(data)) {
        data.forEach(function (f) {
          const key = f.tabId || f.id || f.fieldId || f.key;
          const label = f.tabLabel || f.label || f.name || '';
          if (key && !seen[key]) {
            seen[key] = true;
            out.push({ fieldKey: String(key), fieldLabel: label || String(key) });
          }
        });
      }
      
      // If it's an object with fields/tabs property
      if (data && typeof data === 'object' && !Array.isArray(data)) {
        const fields = data.fields || data.tabs || data.tabList || Object.values(data).find(v => Array.isArray(v));
        if (Array.isArray(fields)) {
          fields.forEach(function (f) {
            const key = f.tabId || f.id || f.fieldId || f.key;
            const label = f.tabLabel || f.label || f.name || '';
            if (key && !seen[key]) {
              seen[key] = true;
              out.push({ fieldKey: String(key), fieldLabel: label || String(key) });
            }
          });
        }
      }
    }
    
    // Try to find React/Redux state in DOM attributes (some frameworks store state here)
    const reactElements = root.querySelectorAll('[data-reactroot], [data-react-helmet], [data-react-class]');
    reactElements.forEach(function (el) {
      try {
        // Try to access React internal state
        const reactKey = Object.keys(el).find(k => k.startsWith('__reactInternalInstance') || k.startsWith('__reactFiber'));
        if (reactKey && el[reactKey]) {
          const fiber = el[reactKey];
          let current = fiber;
          let depth = 0;
          while (current && depth < 20) {
            if (current.memoizedState || current.memoizedProps) {
              const state = current.memoizedState || current.memoizedProps;
              if (state && state.tabs) {
                (Array.isArray(state.tabs) ? state.tabs : [state.tabs]).forEach(function (f) {
                  const key = f.tabId || f.id;
                  if (key && !seen[key]) {
                    seen[key] = true;
                    out.push({ fieldKey: String(key), fieldLabel: (f.tabLabel || f.label || String(key)) });
                  }
                });
              }
            }
            current = current.return || current.child;
            depth++;
          }
        }
      } catch (e) {}
    });
  } catch (e) {
    console.log('DocuFill: Error accessing DocuSign internals', e);
  }
  
  // Strategy 5: Look for clickable elements that might be fields (for canvas/PDF-based rendering)
  const clickableFields = root.querySelectorAll('[onclick*="field"], [onclick*="tab"], [data-clickable="true"], [class*="field"][class*="click"], [class*="tab"][class*="click"]');
  clickableFields.forEach(function (el) {
    const rect = el.getBoundingClientRect();
    // Only consider elements that look like form fields (have reasonable size)
    if (rect.width > 50 && rect.height > 15 && rect.width < 1000 && rect.height < 100) {
      const info = extractFieldInfo(el);
      const key = info.key || buildXPath(el);
      if (key && !seen[key]) {
        seen[key] = true;
        out.push({ fieldKey: key, fieldLabel: (info.label || '').trim() || key });
      }
    }
  });
  
  // Strategy 6: Look for elements with DocuSign-specific class patterns
  const dsClassFields = root.querySelectorAll('[class*="ds-"], [class*="docusign-"], [class*="tab-"], [class*="field-"], [class*="signature-"], [class*="text-field"]');
  dsClassFields.forEach(function (el) {
    const rect = el.getBoundingClientRect();
    // Check if it looks like a field (has size, is visible, might be clickable)
    if (rect.width > 20 && rect.height > 10 && rect.width < 2000) {
      const tabId = el.getAttribute('data-tab-id') || el.getAttribute('id') || el.getAttribute('data-id');
      const label = el.getAttribute('data-tab-label') || el.getAttribute('aria-label') || el.getAttribute('title') || 
                   el.textContent.trim().slice(0, 50);
      if (tabId && !seen[tabId]) {
        seen[tabId] = true;
        out.push({ fieldKey: tabId, fieldLabel: label || tabId });
      } else if (!tabId) {
        const info = extractFieldInfo(el);
        const key = info.key || buildXPath(el);
        if (key && !seen[key] && info.label && info.label !== 'Text' && info.label !== 'Unknown Field') {
          seen[key] = true;
          out.push({ fieldKey: key, fieldLabel: info.label });
        }
      }
    }
  });
  
  return out;
}

function activateCaptureMode() {
  captureMode = true;
  createOverlay();
  
  // Multiple console methods to ensure visibility
  console.log('%c🎯🎯🎯 DOCUFILL CAPTURE MODE ACTIVATED 🎯🎯🎯', 'color: #22c55e; font-size: 18px; font-weight: bold; background: #1a1d27; padding: 8px; border: 2px solid #22c55e; border-radius: 4px;');
  console.warn('⚠️ DocuFill: CAPTURE MODE IS NOW ACTIVE - Click a field to test!');
  console.log('DocuFill: Frame type:', isTopFrame() ? 'top' : 'child');
  console.log('DocuFill: URL:', window.location.href);
  console.log('DocuFill: Now monitoring for field clicks and DocuSign property panel updates...');
  console.log('DocuFill: Click a field in the document to test!');
  
  // Use capture phase to catch events before they're stopped
  document.addEventListener('mouseover', onMouseOver, true);
  document.addEventListener('mouseout', onMouseOut, true);
  document.addEventListener('click', onClick, true);
  // Also listen on window in case events bubble there
  window.addEventListener('click', onClick, true);
  // Listen on document.body as well
  if (document.body) {
    document.body.addEventListener('click', onClick, true);
  }
  
  // Try to inject into iframes
  if (isTopFrame()) {
    const iframes = document.querySelectorAll('iframe');
    console.log('DocuFill: Found', iframes.length, 'iframe(s)');
    
    setTimeout(() => {
      iframes.forEach(function (frame, index) {
        try {
          const src = frame.src || frame.getAttribute('src') || 'no src';
          const frameWindow = frame.contentWindow;
          const frameDoc = frame.contentDocument || (frameWindow && frameWindow.document);
          
          console.log('DocuFill: Attempting to access iframe', index, 'src:', src.substring(0, 100));
          
          // Method 1: Try to access frame directly (same-origin only)
          if (frameDoc) {
            console.log('DocuFill: Iframe', index, 'is same-origin! Injecting directly...');
            // Inject our capture mode into the iframe's document
            try {
              const script = frameDoc.createElement('script');
              script.textContent = `
                (function() {
                  if (window.${DOCUFILL_PREFIX.replace(/[^a-zA-Z0-9]/g, '_')}captureMode) return;
                  window.${DOCUFILL_PREFIX.replace(/[^a-zA-Z0-9]/g, '_')}captureMode = true;
                  document.addEventListener('click', function(e) {
                    console.log('🔵 DocuFill: Click in iframe detected!', e.target);
                    window.parent.postMessage({ type: '${DOCUFILL_PREFIX}IFRAME_CLICK', target: {
                      tagName: e.target.tagName,
                      id: e.target.id,
                      className: e.target.className,
                      textContent: e.target.textContent?.substring(0, 50),
                      x: e.clientX,
                      y: e.clientY
                    }}, '*');
                  }, true);
                })();
              `;
              (frameDoc.head || frameDoc.documentElement).appendChild(script);
              console.log('DocuFill: Successfully injected script into iframe', index);
            } catch (injectError) {
              console.warn('DocuFill: Could not inject script into iframe', index, injectError);
            }
          }
          
          // Method 2: Try to access PDF.js API if it exists
          if (frameWindow) {
            try {
              // Check for PDF.js
              if (frameWindow.PDFViewerApplication || frameWindow.pdfjsLib) {
                console.log('DocuFill: PDF.js detected in iframe', index);
                // Try to get field information from PDF.js
                if (frameWindow.PDFViewerApplication && frameWindow.PDFViewerApplication.pdfDocument) {
                  frameWindow.PDFViewerApplication.pdfDocument.getFieldObjects().then(function(fields) {
                    console.log('DocuFill: Found PDF fields via PDF.js:', fields);
                    if (fields && fields.length > 0) {
                      fields.forEach(function(field) {
                        window.postMessage({ 
                          type: DOCUFILL_PREFIX + 'PDF_FIELD_DETECTED',
                          field: {
                            name: field.name || field.id,
                            type: field.type,
                            value: field.value
                          }
                        }, '*');
                      });
                    }
                  }).catch(function(e) {
                    console.log('DocuFill: PDF.js getFieldObjects not available:', e);
                  });
                }
              }
            } catch (pdfError) {
              // PDF.js not available or different version
            }
          }
          
          // Method 3: Post message (works for both same-origin and cross-origin)
          frameWindow.postMessage({ type: DOCUFILL_PREFIX + 'CAPTURE_MODE_ON' }, '*');
          console.log('DocuFill: Posted message to iframe', index);
          
        } catch (e) {
          // Cross-origin iframe - can't access directly
          console.warn('DocuFill: Cannot access iframe', index, '(likely cross-origin):', e.message);
          console.log('DocuFill: This iframe may contain the PDF viewer. Cross-origin iframes cannot be accessed for security reasons.');
        }
      });
    }, 500);
    
    // Listen for messages from iframes
    window.addEventListener('message', function(ev) {
      if (ev.data && ev.data.type === DOCUFILL_PREFIX + 'IFRAME_CLICK') {
        console.log('DocuFill: Received click from iframe:', ev.data.target);
        // Try to capture this click
        const fieldInfo = {
          key: ev.data.target.id || `iframe_field_${ev.data.target.x}_${ev.data.target.y}`,
          label: ev.data.target.textContent || ev.data.target.className || 'Field from iframe'
        };
        chrome.runtime.sendMessage({ type: 'FIELD_CAPTURED', fieldKey: fieldInfo.key, fieldLabel: fieldInfo.label });
      }
      if (ev.data && ev.data.type === DOCUFILL_PREFIX + 'PDF_FIELD_DETECTED') {
        console.log('DocuFill: PDF field detected:', ev.data.field);
        const fieldInfo = {
          key: ev.data.field.name || ev.data.field.id,
          label: ev.data.field.name || 'PDF Field'
        };
        chrome.runtime.sendMessage({ type: 'FIELD_CAPTURED', fieldKey: fieldInfo.key, fieldLabel: fieldInfo.label });
      }
    });
  }
  
    // Prepare Mode Strategy: Monitor DocuSign's property panel for field selection
    console.log('DocuFill: Setting up prepare mode field detection...');
    
    // Strategy 1: Poll for selected field in DocuSign's property panel every 500ms
    const checkSelectedField = setInterval(() => {
      if (!captureMode) {
        clearInterval(checkSelectedField);
        return;
      }
      
      try {
        // Look for DocuSign's field property panel - try multiple selectors
        const sidebars = [
          document.querySelector('[class*="sidebar"]'),
          document.querySelector('[class*="panel"]'),
          document.querySelector('[class*="properties"]'),
          document.querySelector('[class*="field-properties"]'),
          document.querySelector('[class*="tab-properties"]'),
          // Look for the right-side panel that shows field details
          Array.from(document.querySelectorAll('*')).find(el => {
            const text = el.textContent || '';
            return text.includes('Selected Fields') || text.includes('Add Text') || text.includes('Character Limit');
          })
        ].filter(Boolean);
        
        sidebars.forEach(sidebar => {
          if (!sidebar) return;
          
          // Strategy 1: Look for "X Selected Fields" text - this tells us a field is selected
          const selectedFieldsText = sidebar.textContent?.match(/(\d+)\s+Selected\s+Fields?/i);
          if (selectedFieldsText && parseInt(selectedFieldsText[1]) > 0) {
            // A field is selected - try to find its identifier
            const allInputs = sidebar.querySelectorAll('input, [contenteditable], textarea, select');
            allInputs.forEach(input => {
              const value = (input.value || input.textContent || input.getAttribute('value') || '').trim();
              // If input has a value, it might be a field name/ID
              if (value && value.length > 0 && value !== 'Text' && value.length < 100) {
                // Try to find a label for context
                let label = '';
                // Look for nearby text that might be a label
                let sibling = input.previousElementSibling;
                let attempts = 0;
                while (sibling && attempts < 5) {
                  const text = sibling.textContent?.trim();
                  if (text && text.length < 50 && !text.includes('Character')) {
                    label = text;
                    break;
                  }
                  sibling = sibling.previousElementSibling;
                  attempts++;
                }
                // If no label found, check parent
                if (!label) {
                  const parentText = input.parentElement?.textContent?.trim();
                  if (parentText && parentText.length < 100) {
                    label = parentText.replace(value, '').trim();
                  }
                }
                
                console.warn('⚠️ DocuFill: Found field value in property panel:', value, label);
                chrome.runtime.sendMessage({ 
                  type: 'DOCUSIGN_FIELD_DETECTED', 
                  fieldKey: value, 
                  fieldLabel: label || value 
                });
              }
            });
          }
          
          // Strategy 2: Look for any input with a value (might be field name/ID)
          const inputs = sidebar.querySelectorAll('input[value], [contenteditable], input[type="text"]');
          inputs.forEach(input => {
            const value = (input.value || input.textContent || input.getAttribute('value') || '').trim();
            const label = input.getAttribute('aria-label') || 
                         input.getAttribute('placeholder') ||
                         input.getAttribute('title') ||
                         input.previousElementSibling?.textContent?.trim() ||
                         input.closest('label')?.textContent?.trim() ||
                         input.parentElement?.querySelector('label')?.textContent?.trim();
            
            // More lenient - if it has a value and looks like it could be a field identifier
            if (value && value.length > 0 && value !== 'Text' && value.length < 100 && 
                !value.match(/^\d+$/) && // Not just a number
                value !== '4000') { // Not the character limit
              // Check if this input is in a properties/field context
              const context = input.closest('[class*="properties"], [class*="field"], [class*="tab"]')?.textContent || '';
              if (context.includes('Field') || context.includes('Text') || context.includes('Name') || 
                  label?.toLowerCase().includes('name') || label?.toLowerCase().includes('id') || 
                  label?.toLowerCase().includes('label') || input.id?.toLowerCase().includes('name') ||
                  input.id?.toLowerCase().includes('id') || input.name?.toLowerCase().includes('name')) {
                console.warn('⚠️ DocuFill: Found potential field in property panel:', value, label);
                chrome.runtime.sendMessage({ 
                  type: 'DOCUSIGN_FIELD_DETECTED', 
                  fieldKey: value, 
                  fieldLabel: label || value 
                });
              }
            }
          });
        });
        
        // Strategy 2: Look for selected/highlighted field elements
        const selectedFields = document.querySelectorAll('[class*="selected"][data-tab-id], [class*="active"][data-tab-id], [aria-selected="true"][data-tab-id]');
        selectedFields.forEach(el => {
          const tabId = el.getAttribute('data-tab-id') || el.getAttribute('data-field-id');
          if (tabId) {
            const label = el.getAttribute('data-tab-label') ||
                         el.getAttribute('aria-label') ||
                         el.textContent?.trim();
            console.log('DocuFill: Found selected field:', tabId, label);
            chrome.runtime.sendMessage({ 
              type: 'DOCUSIGN_FIELD_DETECTED', 
              fieldKey: String(tabId), 
              fieldLabel: label || String(tabId) 
            });
          }
        });
      } catch (e) {
        // Silent fail - don't spam console
      }
    }, 500); // Check every 500ms
    
    window.__docufill_fieldChecker = checkSelectedField;
    
    // Strategy 3: Intercept clicks and read DocuSign's response after a delay
    const prepareModeClickHandler = function(e) {
      if (!captureMode) return;
      
      // After DocuSign processes the click, check for field selection
      setTimeout(() => {
        try {
          // Look for recently updated/selected elements
          const recentlySelected = document.querySelector('[class*="selected"][data-tab-id], [class*="active"][data-tab-id]');
          if (recentlySelected) {
            const tabId = recentlySelected.getAttribute('data-tab-id');
            const label = recentlySelected.getAttribute('data-tab-label') ||
                         recentlySelected.getAttribute('aria-label') ||
                         recentlySelected.textContent?.trim();
            
            if (tabId) {
              console.log('%c🔵 DocuFill: Detected field selection after click:', 'color: #22c55e; font-weight: bold;', tabId, label);
              chrome.runtime.sendMessage({ 
                type: 'DOCUSIGN_FIELD_DETECTED', 
                fieldKey: String(tabId), 
                fieldLabel: label || String(tabId) 
              });
            }
          }
          
          // Also check property panel for current field
          const propertyPanel = document.querySelector('[class*="properties"], [class*="sidebar"]');
          if (propertyPanel) {
            const nameInput = propertyPanel.querySelector('input[value], [contenteditable]');
            if (nameInput) {
              const fieldName = nameInput.value || nameInput.textContent;
              if (fieldName && fieldName.trim().length > 0) {
                console.log('DocuFill: Found field name in property panel:', fieldName);
                chrome.runtime.sendMessage({ 
                  type: 'DOCUSIGN_FIELD_DETECTED', 
                  fieldKey: String(fieldName).trim(), 
                  fieldLabel: fieldName.trim() 
                });
              }
            }
          }
        } catch (err) {
          console.warn('DocuFill: Error in prepare mode click handler', err);
        }
      }, 300); // Wait for DocuSign to update
    };
    
    document.addEventListener('click', prepareModeClickHandler, true);
    window.__docufill_prepareClickHandler = prepareModeClickHandler;
    
    // Try to find and click on DocuSign field elements directly
    setTimeout(() => {
      console.log('DocuFill: Scanning for DocuSign field elements...');
      const possibleFields = document.querySelectorAll('[data-tab-id], [data-field-id], [class*="field"], [class*="tab"], [role="textbox"]');
      console.log('DocuFill: Found', possibleFields.length, 'potential field elements');
      if (possibleFields.length > 0) {
        console.log('DocuFill: Sample field elements:', Array.from(possibleFields.slice(0, 3)).map(el => ({
          tag: el.tagName,
          id: el.id,
          class: el.className,
          dataTabId: el.getAttribute('data-tab-id'),
          dataFieldId: el.getAttribute('data-field-id')
        })));
      }
      
      // Try to extract fields from DocuSign's internal data
      extractDocuSignInternalFields();
    }, 1000);
}

// Function to manually read fields from DocuSign's property panel
function readPropertyPanelFields() {
  console.log('DocuFill: Manually reading property panel...');
  const fields = [];
  const seen = {};
  
  try {
    // Find the right sidebar/property panel - look for "Selected Fields" text
    const allElements = Array.from(document.querySelectorAll('*'));
    const sidebars = [];
    
    // Find elements containing "Selected Fields" or property panel indicators
    for (let el of allElements) {
      const text = el.textContent || '';
      if (text.includes('Selected Fields') || text.includes('Add Text') || 
          text.includes('Character Limit') || text.includes('Required Field')) {
        sidebars.push(el);
      }
    }
    
    // Also check common sidebar/panel selectors
    sidebars.push(...Array.from(document.querySelectorAll('[class*="sidebar"]')));
    sidebars.push(...Array.from(document.querySelectorAll('[class*="panel"]')));
    sidebars.push(...Array.from(document.querySelectorAll('[class*="properties"]')));
    
    // Remove duplicates
    const uniqueSidebars = [...new Set(sidebars)];
    
    console.log('DocuFill: Found', uniqueSidebars.length, 'potential property panel(s)');
    
    // DEBUG: Log all sidebar text content to see what we're working with
    uniqueSidebars.forEach((sidebar, idx) => {
      const text = sidebar.textContent || '';
      if (text.includes('Selected') || text.includes('Field')) {
        console.log(`DocuFill: Sidebar ${idx} text (first 200 chars):`, text.substring(0, 200));
        console.log(`DocuFill: Sidebar ${idx} HTML (first 500 chars):`, sidebar.innerHTML?.substring(0, 500));
      }
    });
    
    uniqueSidebars.forEach(sidebar => {
      // Look for "X Selected Fields" - indicates a field is selected
      const selectedText = sidebar.textContent?.match(/(\d+)\s+Selected\s+Fields?/i);
      if (selectedText && parseInt(selectedText[1]) > 0) {
        console.log('DocuFill: Found "Selected Fields" indicator -', selectedText[1], 'field(s) selected!');
        
        // Find all inputs in this sidebar
        const inputs = sidebar.querySelectorAll('input, [contenteditable], textarea, [role="textbox"]');
        console.log('DocuFill: Found', inputs.length, 'input(s) in property panel');
        
        // DEBUG: Log all input values
        inputs.forEach((input, index) => {
          const value = (input.value || input.textContent || input.getAttribute('value') || '').trim();
          const type = input.type || input.tagName;
          const id = input.id || input.getAttribute('data-testid') || 'no-id';
          console.log(`DocuFill: Input ${index}: type=${type}, id=${id}, value="${value}"`);
        });
        
        inputs.forEach((input, index) => {
          const value = (input.value || input.textContent || '').trim();
          if (value && value.length > 0 && value !== 'Text' && value.length < 100 && value !== '4000') {
            // Try to find label
            let label = input.getAttribute('aria-label') || 
                       input.getAttribute('placeholder') ||
                       input.getAttribute('title') ||
                       input.getAttribute('name');
            
            // Look for nearby text that might be a label
            let sibling = input.previousElementSibling;
            for (let i = 0; i < 5 && !label; i++) {
              if (sibling) {
                const text = sibling.textContent?.trim();
                if (text && text.length < 50 && !text.match(/^\d+$/) && 
                    !text.includes('Character') && !text.includes('Limit')) {
                  label = text;
                }
                sibling = sibling.previousElementSibling;
              }
            }
            
            // Check parent for context
            if (!label) {
              const parent = input.parentElement;
              if (parent) {
                const parentText = parent.textContent?.trim();
                const parts = parentText.split(value);
                if (parts[0]) {
                  label = parts[0].trim().slice(-40);
                }
              }
            }
            
            // Also try to find label by looking for text nodes near the input
            if (!label) {
              const walker = document.createTreeWalker(
                input.parentElement || sidebar,
                NodeFilter.SHOW_TEXT,
                null
              );
              let node;
              while (node = walker.nextNode()) {
                const text = node.textContent?.trim();
                if (text && text.length < 50 && text !== value && 
                    !text.match(/^\d+$/) && !text.includes('Character') && 
                    !text.includes('Limit') && !text.includes('Selected')) {
                  // Check if this text node is near our input
                  const rect = input.getBoundingClientRect();
                  const textRect = node.parentElement?.getBoundingClientRect();
                  if (textRect && Math.abs(rect.top - textRect.top) < 50) {
                    label = text;
                    break;
                  }
                }
              }
            }
            
            if (value && !seen[value]) {
              seen[value] = true;
              fields.push({ fieldKey: value, fieldLabel: label || value });
              console.log('DocuFill: Extracted field from property panel:', value, label || '(no label)');
            }
          }
        });
        
        // Also try to find field identifiers in data attributes
        const elementsWithData = sidebar.querySelectorAll('[data-*]');
        elementsWithData.forEach(el => {
          Array.from(el.attributes).forEach(attr => {
            if (attr.name.startsWith('data-') && attr.value && attr.value.length > 0 && attr.value.length < 100) {
              const attrName = attr.name.replace('data-', '');
              if (attrName.includes('id') || attrName.includes('name') || attrName.includes('key') || 
                  attrName.includes('field') || attrName.includes('tab')) {
                if (!seen[attr.value]) {
                  seen[attr.value] = true;
                  fields.push({ fieldKey: attr.value, fieldLabel: attrName + ': ' + attr.value });
                  console.log('DocuFill: Found field in data attribute:', attrName, '=', attr.value);
                }
              }
            }
          });
        });
      }
    });
    
    console.log('DocuFill: Read', fields.length, 'field(s) from property panel');
    
    // If no fields found, return debug info
    if (fields.length === 0) {
      console.warn('DocuFill: No fields found. Property panel structure may be different than expected.');
      console.warn('DocuFill: Try inspecting the right sidebar manually in DevTools to see its structure.');
    }
    
    return fields;
  } catch (e) {
    console.error('DocuFill: Error reading property panel', e);
    return [];
  }
}

// Function to extract field data from DocuSign's internal structures
function extractDocuSignInternalFields() {
  console.log('DocuFill: Attempting to extract fields from DocuSign internals...');
  const fields = [];
  
  // Strategy for Prepare Mode: Try to read from DocuSign's UI sidebar
  try {
    // Look for DocuSign's field property panel/sidebar
    const sidebarSelectors = [
      '[class*="field-properties"]',
      '[class*="field-panel"]',
      '[class*="tab-properties"]',
      '[class*="properties-panel"]',
      '[data-testid*="field"]',
      '[data-testid*="properties"]',
      '[aria-label*="field"]',
      '[aria-label*="properties"]'
    ];
    
    sidebarSelectors.forEach(selector => {
      try {
        const panels = document.querySelectorAll(selector);
        panels.forEach(panel => {
          // Try to extract field info from property panels
          const inputs = panel.querySelectorAll('input, select, [contenteditable]');
          inputs.forEach(input => {
            const label = input.getAttribute('aria-label') || 
                         input.getAttribute('placeholder') ||
                         input.previousElementSibling?.textContent ||
                         input.closest('label')?.textContent;
            const value = input.value || input.textContent;
            if (label && value && (label.toLowerCase().includes('name') || label.toLowerCase().includes('label'))) {
              const key = value || `field_${Date.now()}_${Math.random()}`;
              if (!fields.find(f => f.fieldKey === key)) {
                fields.push({ fieldKey: key, fieldLabel: label || value });
              }
            }
          });
        });
      } catch (e) {}
    });
  } catch (e) {
    console.warn('DocuFill: Error reading from sidebar', e);
  }
  
  // Strategy: Hook into DocuSign's field selection events
  try {
    // Override/addEventListener to catch DocuSign's field selection
    const originalAddEventListener = EventTarget.prototype.addEventListener;
    EventTarget.prototype.addEventListener = function(type, listener, options) {
      if (type === 'click' || type === 'select' || type.includes('field') || type.includes('tab')) {
        const wrappedListener = function(event) {
          try {
            // Check if this is a field selection event
            const target = event.target || event.currentTarget;
            if (target) {
              // Look for field indicators in the clicked element
              const tabId = target.getAttribute('data-tab-id') || 
                           target.closest('[data-tab-id]')?.getAttribute('data-tab-id');
              const fieldId = target.getAttribute('data-field-id') ||
                            target.closest('[data-field-id]')?.getAttribute('data-field-id');
              
              if (tabId || fieldId) {
                const key = tabId || fieldId;
                const label = target.getAttribute('data-tab-label') ||
                             target.getAttribute('aria-label') ||
                             target.textContent?.trim() ||
                             target.closest('[data-tab-label]')?.getAttribute('data-tab-label');
                
                if (key && !fields.find(f => f.fieldKey === key)) {
                  fields.push({ fieldKey: String(key), fieldLabel: label || String(key) });
                  chrome.runtime.sendMessage({ 
                    type: 'DOCUSIGN_FIELD_DETECTED', 
                    fieldKey: String(key), 
                    fieldLabel: label || String(key) 
                  });
                }
              }
            }
          } catch (e) {}
          return listener.apply(this, arguments);
        };
        return originalAddEventListener.call(this, type, wrappedListener, options);
      }
      return originalAddEventListener.call(this, type, listener, options);
    };
  } catch (e) {
    console.warn('DocuFill: Could not hook event system', e);
  }
  
  // Strategy: Monitor DOM mutations for field creation
  try {
    const observer = new MutationObserver(function(mutations) {
      mutations.forEach(function(mutation) {
        mutation.addedNodes.forEach(function(node) {
          if (node.nodeType === 1) { // Element node
            const tabId = node.getAttribute && node.getAttribute('data-tab-id');
            const fieldId = node.getAttribute && node.getAttribute('data-field-id');
            if (tabId || fieldId) {
              const key = tabId || fieldId;
              const label = node.getAttribute && (
                node.getAttribute('data-tab-label') ||
                node.getAttribute('aria-label') ||
                node.textContent?.trim()
              );
              if (key && !fields.find(f => f.fieldKey === key)) {
                fields.push({ fieldKey: String(key), fieldLabel: label || String(key) });
                chrome.runtime.sendMessage({ 
                  type: 'DOCUSIGN_FIELD_DETECTED', 
                  fieldKey: String(key), 
                  fieldLabel: label || String(key) 
                });
              }
            }
          }
        });
      });
    });
    
    observer.observe(document.body || document.documentElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['data-tab-id', 'data-field-id', 'data-tab-label']
    });
    
    // Store observer so it doesn't get garbage collected
    window.__docufill_observer = observer;
  } catch (e) {
    console.warn('DocuFill: Could not set up mutation observer', e);
  }
  
  // Strategy: Try to access DocuSign's field management API
  try {
    // Look for DocuSign's field manager or editor
    if (window.DocuSign && window.DocuSign.Editor) {
      try {
        const editor = window.DocuSign.Editor;
        if (editor.getTabs && typeof editor.getTabs === 'function') {
          const tabs = editor.getTabs();
          if (Array.isArray(tabs)) {
            tabs.forEach(tab => {
              const key = tab.tabId || tab.id;
              if (key && !fields.find(f => f.fieldKey === key)) {
                fields.push({ fieldKey: String(key), fieldLabel: (tab.tabLabel || tab.label || String(key)) });
              }
            });
          }
        }
      } catch (e) {}
    }
  } catch (e) {
    console.warn('DocuFill: Error accessing DocuSign Editor API', e);
  }
  
  try {
    // Method 1: Try Redux store
    const reduxStores = [];
    
    // Look for Redux store in common locations
    if (window.__REDUX_STORE__) reduxStores.push(window.__REDUX_STORE__);
    if (window.store) reduxStores.push(window.store);
    if (window.__store__) reduxStores.push(window.__store__);
    
    // Try to find Redux store by checking window properties
    for (let key in window) {
      try {
        const obj = window[key];
        if (obj && typeof obj.getState === 'function' && typeof obj.dispatch === 'function') {
          reduxStores.push(obj);
        }
      } catch (e) {}
    }
    
    reduxStores.forEach((store, index) => {
      try {
        const state = store.getState();
        console.log('DocuFill: Found Redux store', index, 'with state keys:', Object.keys(state || {}));
        
        // Common Redux state paths for DocuSign
        const possiblePaths = [
          state?.tabs,
          state?.fields,
          state?.document?.tabs,
          state?.envelope?.tabs,
          state?.prepare?.tabs,
          state?.signing?.tabs,
          state?.document?.fields,
          state?.envelope?.fields,
          state?.prepare?.fields,
          state?.signing?.fields,
          state?.document?.recipients?.[0]?.tabs,
          state?.envelope?.recipients?.[0]?.tabs,
          state?.prepare?.recipients?.[0]?.tabs,
          state?.signing?.recipients?.[0]?.tabs,
          state?.ui?.document?.tabs,
          state?.editor?.tabs,
          state?.editor?.fields
        ];
        
        possiblePaths.forEach((tabs, pathIndex) => {
          if (Array.isArray(tabs) && tabs.length > 0) {
            console.log('DocuFill: Found tabs array at path', pathIndex, 'with', tabs.length, 'tabs');
            tabs.forEach((tab) => {
              const key = tab.tabId || tab.id || tab.fieldId || tab.key || 
                          `${tab.recipientId || 'r'}_${tab.documentId || 'd'}_${tab.pageNumber || 0}_${tab.xPosition || 0}_${tab.yPosition || 0}`;
              const label = tab.tabLabel || tab.label || tab.name || tab.type || '';
              if (key && !fields.find(f => f.fieldKey === key)) {
                fields.push({ fieldKey: String(key), fieldLabel: label || String(key) });
              }
            });
          } else if (tabs && typeof tabs === 'object' && !Array.isArray(tabs)) {
            // Might be an object with tab arrays as values
            Object.values(tabs).forEach((value) => {
              if (Array.isArray(value)) {
                value.forEach((tab) => {
                  const key = tab.tabId || tab.id || tab.fieldId || tab.key;
                  const label = tab.tabLabel || tab.label || tab.name || '';
                  if (key && !fields.find(f => f.fieldKey === key)) {
                    fields.push({ fieldKey: String(key), fieldLabel: label || String(key) });
                  }
                });
              }
            });
          }
        });
      } catch (e) {
        console.warn('DocuFill: Error reading Redux store', index, e);
      }
    });
    
    // Method 2: Try window objects that might contain field data
    const windowKeys = Object.keys(window).filter(k => 
      k.toLowerCase().includes('docusign') || 
      k.toLowerCase().includes('ds') ||
      k.toLowerCase().includes('tab') ||
      k.toLowerCase().includes('field') ||
      k.toLowerCase().includes('envelope') ||
      k.toLowerCase().includes('document')
    );
    
    windowKeys.forEach(key => {
      try {
        const obj = window[key];
        if (obj && typeof obj === 'object') {
          // Check if it has tabs or fields
          if (obj.tabs && Array.isArray(obj.tabs)) {
            obj.tabs.forEach(tab => {
              const key = tab.tabId || tab.id;
              if (key && !fields.find(f => f.fieldKey === key)) {
                fields.push({ fieldKey: String(key), fieldLabel: (tab.tabLabel || tab.label || String(key)) });
              }
            });
          }
          if (obj.fields && Array.isArray(obj.fields)) {
            obj.fields.forEach(field => {
              const key = field.id || field.fieldId;
              if (key && !fields.find(f => f.fieldKey === key)) {
                fields.push({ fieldKey: String(key), fieldLabel: (field.label || field.name || String(key)) });
              }
            });
          }
        }
      } catch (e) {}
    });
    
    // Method 3: Try to access DocuSign's API if exposed
    if (window.DocuSign) {
      try {
        // Check various DocuSign API objects
        const dsObjects = [
          window.DocuSign.tabs,
          window.DocuSign.fields,
          window.DocuSign.envelope?.tabs,
          window.DocuSign.document?.tabs,
          window.DocuSign.currentEnvelope?.tabs,
          window.DocuSign.currentDocument?.tabs
        ];
        
        dsObjects.forEach((obj, index) => {
          if (Array.isArray(obj)) {
            obj.forEach(item => {
              const key = item.tabId || item.id;
              if (key && !fields.find(f => f.fieldKey === key)) {
                fields.push({ fieldKey: String(key), fieldLabel: (item.tabLabel || item.label || String(key)) });
              }
            });
          }
        });
      } catch (e) {
        console.warn('DocuFill: Error accessing DocuSign API', e);
      }
    }
    
    console.log('DocuFill: Extracted', fields.length, 'fields from DocuSign internals');
    if (fields.length > 0) {
      console.log('DocuFill: Sample fields:', fields.slice(0, 5));
      // Send fields to popup
      fields.forEach(field => {
        chrome.runtime.sendMessage({ 
          type: 'DOCUSIGN_FIELD_DETECTED', 
          fieldKey: field.fieldKey, 
          fieldLabel: field.fieldLabel 
        });
      });
    }
    
    return fields;
  } catch (e) {
    console.error('DocuFill: Error extracting DocuSign fields', e);
    return [];
  }
}

function deactivateCaptureMode() {
  captureMode = false;
  removeOverlay();
  clearHighlight();
  document.removeEventListener('mouseover', onMouseOver, true);
  document.removeEventListener('mouseout', onMouseOut, true);
  document.removeEventListener('click', onClick, true);
  window.removeEventListener('click', onClick, true);
  
  // Clean up prepare mode handlers
  if (window.__docufill_fieldChecker) {
    clearInterval(window.__docufill_fieldChecker);
    window.__docufill_fieldChecker = null;
  }
  if (window.__docufill_prepareClickHandler) {
    document.removeEventListener('click', window.__docufill_prepareClickHandler, true);
    window.__docufill_prepareClickHandler = null;
  }
  if (window.__docufill_observer) {
    window.__docufill_observer.disconnect();
    window.__docufill_observer = null;
  }
}

function resolveClickTarget(target) {
  if (!target || !target.getBoundingClientRect) return null;
  
  // If it's a label, try to find the associated input
  if (target.tagName === 'LABEL' && target.htmlFor) {
    const forEl = document.getElementById(target.htmlFor);
    if (forEl) {
      const input = findInputElement(forEl) || forEl;
      if (input) return input;
    }
  }
  
  // First, try to find an actual input element
  const inputEl = findInputElement(target);
  if (inputEl) return inputEl;
  
  // If no input found, check if the clicked element itself might be a field
  // (for DocuSign prepare mode where fields are visual placeholders)
  const rect = target.getBoundingClientRect();
  const hasReasonableSize = rect.width > 20 && rect.height > 10 && rect.width < 2000 && rect.height < 500;
  
  // Check for DocuSign field indicators
  const hasFieldIndicator = 
    target.getAttribute('data-tab-id') ||
    target.getAttribute('data-tab-label') ||
    target.getAttribute('data-field-id') ||
    target.classList.contains('ds-field') ||
    target.getAttribute('data-testid')?.includes('field') ||
    target.getAttribute('role') === 'textbox' ||
    target.getAttribute('tabindex') !== null ||
    (target.className && (target.className.includes('field') || target.className.includes('tab'))) ||
    (hasReasonableSize && (target.onclick || target.style.cursor === 'pointer' || target.getAttribute('onclick')));
  
  // If it looks like a field and has reasonable size, capture it
  if (hasFieldIndicator && hasReasonableSize) {
    return target;
  }
  
  // Walk up the DOM tree to find a parent that might be a field
  let current = target.parentElement;
  let depth = 0;
  while (current && depth < 5) {
    const parentInput = findInputElement(current);
    if (parentInput) return parentInput;
    
    const parentHasFieldIndicator = 
      current.getAttribute('data-tab-id') ||
      current.getAttribute('data-field-id') ||
      current.classList.contains('ds-field');
    
    if (parentHasFieldIndicator && hasReasonableSize) {
      return current;
    }
    
    current = current.parentElement;
    depth++;
  }
  
  return null;
}

function onMouseOver(e) {
  if (!captureMode) return;
  const el = resolveClickTarget(e.target);
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
  
  // Log for debugging - this should appear in console
  console.log('🔵 DocuFill: Click detected!', {
    target: e.target,
    tagName: e.target?.tagName,
    className: e.target?.className?.substring(0, 50),
    id: e.target?.id,
    hasDataTabId: !!e.target?.getAttribute('data-tab-id'),
    hasDataFieldId: !!e.target?.getAttribute('data-field-id'),
    x: e.clientX,
    y: e.clientY,
    frame: isTopFrame() ? 'top' : 'child'
  });
  
  // Try to find a field element
  let el = resolveClickTarget(e.target);
  
  // If no field found, try the target itself or walk up the tree
  if (!el) {
    let target = e.target;
    let depth = 0;
    while (target && depth < 10) {
      const rect = target.getBoundingClientRect && target.getBoundingClientRect();
      if (rect && rect.width > 15 && rect.height > 10 && rect.width < 2000 && rect.height < 500) {
        const style = window.getComputedStyle(target);
        if (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0') {
          // This element looks like it could be a field - try to extract info
          const fieldInfo = extractFieldInfo(target);
          if (fieldInfo.key) {
            el = target;
            break;
          }
        }
      }
      target = target.parentElement;
      depth++;
    }
  }
  
  // If we still don't have an element, capture the click anyway with a generated key
  if (!el && e.target) {
    const target = e.target;
    const rect = target.getBoundingClientRect && target.getBoundingClientRect();
    if (rect && rect.width > 15 && rect.height > 10) {
      // Generate a key based on position and nearby text
      const x = Math.round(rect.left);
      const y = Math.round(rect.top);
      const key = `field_${x}_${y}_${Date.now()}`;
      let label = '';
      
      // Try to find nearby text that might be a label
      const nearbyText = findNearbyText(target);
      if (nearbyText) label = nearbyText;
      else label = `Field at (${x}, ${y})`;
      
      el = target;
      const fieldInfo = { key, label };
      
      e.preventDefault();
      e.stopPropagation();
      chrome.runtime.sendMessage({ type: 'FIELD_CAPTURED', fieldKey: fieldInfo.key, fieldLabel: fieldInfo.label });
      target.style.outline = '2px solid #22c55e';
      setTimeout(() => { target.style.outline = ''; }, 800);
      return;
    }
  }
  
  if (!el) return;
  
  e.preventDefault();
  e.stopPropagation();
  const fieldInfo = extractFieldInfo(el);
  chrome.runtime.sendMessage({ type: 'FIELD_CAPTURED', fieldKey: fieldInfo.key, fieldLabel: fieldInfo.label });
  clearHighlight();
  el.style.outline = '2px solid #22c55e';
  setTimeout(() => { el.style.outline = el.__docufill_origOutline || ''; }, 800);
}

function findNearbyText(el) {
  if (!el) return '';
  // Look for text in the element or nearby elements
  const text = el.textContent || el.innerText || '';
  if (text && text.trim().length > 0 && text.trim().length < 100) {
    return text.trim().slice(0, 50);
  }
  // Check parent
  if (el.parentElement) {
    const parentText = el.parentElement.textContent || '';
    if (parentText && parentText.trim().length > 0 && parentText.trim().length < 100) {
      return parentText.trim().slice(0, 50);
    }
  }
  // Check previous sibling (often the label)
  let sibling = el.previousElementSibling;
  if (sibling) {
    const siblingText = sibling.textContent || '';
    if (siblingText && siblingText.trim().length > 0 && siblingText.trim().length < 100) {
      return siblingText.trim().slice(0, 50);
    }
  }
  return '';
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
    let el = findFieldByKey(mapping.fieldKey);
    if (!el && mapping.fieldLabel && String(mapping.fieldLabel).trim() && String(mapping.fieldLabel) !== String(mapping.fieldKey)) {
      el = findFieldByKey(mapping.fieldLabel);
    }
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
  } else if (el.isContentEditable || el.getAttribute('contenteditable') === 'true') {
    el.focus();
    el.textContent = value;
    triggerEvents(el, ['input', 'change', 'blur', 'keyup']);
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

function findInputInTree(root) {
  if (!root) return null;
  const tag = root.tagName?.toLowerCase();
  if (['input', 'select', 'textarea'].includes(tag)) {
    if (tag === 'input' && ['hidden', 'submit', 'button', 'image'].includes((root.type || '').toLowerCase())) return null;
    return root;
  }
  if (root.contentEditable === 'true' || root.getAttribute('contenteditable') === 'true') return root;
  if (root.classList && (root.classList.contains('ds-field') || root.getAttribute('data-testid')?.includes('field') || root.getAttribute('role') === 'textbox' || root.getAttribute('role') === 'combobox')) return root;
  const sel = 'input:not([type="hidden"]):not([type="submit"]):not([type="button"]), select, textarea, [contenteditable="true"], [role="textbox"]';
  const direct = root.querySelector && root.querySelector(sel);
  if (direct) return direct;
  if (root.shadowRoot) {
    const inShadow = findInputInTree(root.shadowRoot) || root.shadowRoot.querySelector(sel);
    if (inShadow) return inShadow;
  }
  const children = root.querySelectorAll && root.querySelectorAll('*');
  if (children) for (let i = 0; i < Math.min(children.length, 20); i++) {
    if (children[i].shadowRoot) {
      const inShadow = findInputInTree(children[i].shadowRoot) || children[i].shadowRoot.querySelector(sel);
      if (inShadow) return inShadow;
    }
  }
  return null;
}

function findInputElement(el) {
  if (!el || !el.getBoundingClientRect) return null;
  let target = el;
  let depth = 0;
  while (target && depth < 10) {
    const found = findInputInTree(target);
    if (found && found !== target) return found;
    const tag = target.tagName?.toLowerCase();
    if (['input', 'select', 'textarea'].includes(tag)) {
      if (tag === 'input' && ['hidden', 'submit', 'button', 'image'].includes((target.type || '').toLowerCase())) { target = target.parentElement; depth++; continue; }
      return target;
    }
    if (target.contentEditable === 'true' || target.getAttribute('contenteditable') === 'true') return target;
    if (target.classList && (target.classList.contains('ds-field') || target.getAttribute('data-testid')?.includes('field') || target.getAttribute('role') === 'textbox' || target.getAttribute('role') === 'combobox')) return target;
    var inner = findInputInTree(target);
    if (inner) return inner;
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
  const dataTabLabel = el.getAttribute('data-tab-label') || el.getAttribute('data-label');
  const title = el.getAttribute('title');
  let label = '';
  if (id) { const labelEl = document.querySelector(`label[for="${id}"]`); if (labelEl) label = labelEl.textContent.trim(); }
  if (!label && dataTabLabel) label = dataTabLabel;
  if (!label && ariaLabel) label = ariaLabel;
  if (!label && title) label = title;
  if (!label && el.placeholder) label = el.placeholder;
  if (!label && name) label = name;
  if (!label && dataId && /[a-zA-Z]/.test(dataId)) label = dataId.replace(/[-_]([a-z])/gi, ' $1').replace(/^./, function (c) { return c.toUpperCase(); });
  if (!label) label = getLabelFromNearbyElement(el);
  if (!label) label = el.className.split(' ')[0] || 'Unknown Field';
  if (label === 'Text' || label === 'Unknown Field') {
    const hint = (name || dataId || id || '').toString();
    if (hint && /name|email|phone|address|license|npi|number|date|sign/i.test(hint)) label = hint.replace(/[-_]([a-z])/gi, ' $1').replace(/^./, function (c) { return c.toUpperCase(); });
  }
  const key = id || name || dataId || buildXPath(el);
  return { key, label };
}

function getLabelFromNearbyElement(el) {
  const trimLabel = (t) => (t && t.length < 80 && !/^[\d\s]+$/.test(t) ? t.replace(/\s+/g, ' ').slice(0, 60) : '');
  let prev = el.previousElementSibling;
  if (prev) {
    const t = (prev.textContent || '').trim();
    const out = trimLabel(t);
    if (out) return out;
  }
  let parent = el.parentElement;
  for (let i = 0; i < 4 && parent; i++) {
    const prevSib = parent.previousElementSibling;
    if (prevSib) {
      const t = (prevSib.textContent || '').trim();
      const out = trimLabel(t);
      if (out) return out;
    }
    const firstChild = parent.firstElementChild;
    if (firstChild && firstChild !== el) {
      const t = (firstChild.textContent || '').trim();
      const out = trimLabel(t);
      if (out) return out;
    }
    parent = parent.parentElement;
  }
  return '';
}

function escapeSelectorAttr(val) {
  return String(val).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

function findInRootByKey(root, key) {
  if (!root) return null;
  let el = null;
  try { if (root.getElementById) el = root.getElementById(key); } catch (e) {}
  if (el) return el;
  const esc = escapeSelectorAttr(key);
  el = root.querySelector && root.querySelector('[name="' + esc + '"]');
  if (el) return el;
  el = root.querySelector && (root.querySelector('[data-id="' + esc + '"]') || root.querySelector('[data-field-id="' + esc + '"]') || root.querySelector('[data-tab-id="' + esc + '"]') || root.querySelector('[data-testid="' + esc + '"]'));
  if (el) return el;
  if (root.evaluate) {
    try {
      const result = root.evaluate(key, root, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
      if (result.singleNodeValue) return result.singleNodeValue;
    } catch (e) {}
  }
  const all = root.querySelectorAll && root.querySelectorAll('*');
  if (all) for (let i = 0; i < all.length; i++) {
    if (all[i].shadowRoot) {
      const inShadow = findInRootByKey(all[i].shadowRoot, key);
      if (inShadow) return inShadow;
    }
  }
  return null;
}

function findFieldByKey(key) {
  let el = findInRootByKey(document, key);
  if (el) return el;
  const walkShadow = (root) => {
    const all = root.querySelectorAll && root.querySelectorAll('*');
    if (all) for (let i = 0; i < all.length; i++) {
      if (all[i].shadowRoot) {
        const found = findInRootByKey(all[i].shadowRoot, key);
        if (found) return found;
        const deep = walkShadow(all[i].shadowRoot);
        if (deep) return deep;
      }
    }
    return null;
  };
  return walkShadow(document);
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
