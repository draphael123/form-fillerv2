// ═══════════════════════════════════════════════════════
//  DocuFill – background.js
//  Service worker: XLSX parsing via SheetJS CDN
// ═══════════════════════════════════════════════════════

importScripts('lib/xlsx.full.min.js');

function getDuplicateColumns(arr) {
  const seen = {};
  const dupes = [];
  arr.forEach(function (c) {
    if (seen[c]) { if (seen[c] === 1) dupes.push(c); }
    else seen[c] = 1;
  });
  return dupes.length ? dupes : null;
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'PARSE_XLSX') {
    try {
      const data = new Uint8Array(message.data);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetNames = workbook.SheetNames || [];
      const sheetIndex = message.sheetIndex != null ? message.sheetIndex : 0;
      const sheetName = sheetNames[sheetIndex] || sheetNames[0];
      if (!sheetName) { sendResponse({ success: false, error: 'No sheets found' }); return true; }
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (message.sheetIndex == null && sheetNames.length > 1) {
        sendResponse({ success: true, sheetNames: sheetNames, needSheetChoice: true });
        return true;
      }
      const columns = jsonData.length > 0 ? Object.keys(jsonData[0]) : [];
      if (jsonData.length === 0 && columns.length === 0) {
        sendResponse({ success: false, error: 'Sheet is empty' });
        return true;
      }
      const duplicateColumns = columns.length ? getDuplicateColumns(columns) : null;
      sendResponse({
        success: true,
        result: { columns, rows: jsonData, duplicateColumns: duplicateColumns }
      });
    } catch (err) {
      sendResponse({ success: false, error: err.message });
    }
    return true;
  }
});
