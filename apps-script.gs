// ════════════════════════════════════════════════════════════════════════════
//  FD Tracker — Google Apps Script Backend
//  Paste this entire file into Extensions → Apps Script in your Google Sheet
//  then Deploy → New Deployment → Web App (Anyone can access)
// ════════════════════════════════════════════════════════════════════════════

const SHEET_NAME = 'FD_Data';
const HEADERS = [
  'id', 'bank', 'fdNum', 'principal', 'rate',
  'startDate', 'maturityDate', 'compounding', 'notes',
  'renewedFromId', 'renewedToId', 'createdAt'
];

// ── Helpers ──────────────────────────────────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setFontWeight('bold')
         .setBackground('#2563eb')
         .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToObject(row) {
  const obj = {};
  HEADERS.forEach((h, i) => obj[h] = row[i] === undefined ? '' : String(row[i]));
  return obj;
}

function objectToRow(obj) {
  return HEADERS.map(h => obj[h] !== undefined ? obj[h] : '');
}

function getAllRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];           // only header
  return data.slice(1).map(rowToObject);
}

function findRowIndex(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based sheet row
  }
  return -1;
}

function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Router ───────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || 'getAll';
    const sheet  = getSheet();

    if (action === 'getAll') {
      return corsResponse({ ok: true, data: getAllRows(sheet) });
    }

    return corsResponse({ ok: false, error: 'Unknown action' });
  } catch(err) {
    return corsResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    const sheet   = getSheet();

    // ── Save all (bulk sync) ──────────────────────────────────────────────
    if (action === 'saveAll') {
      // Clear existing data rows (keep header)
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

      // Write all FDs
      if (payload.data && payload.data.length > 0) {
        const rows = payload.data.map(objectToRow);
        sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
      }
      autoResizeColumns(sheet);
      return corsResponse({ ok: true, count: (payload.data || []).length });
    }

    // ── Add single FD ────────────────────────────────────────────────────
    if (action === 'add') {
      sheet.appendRow(objectToRow(payload.fd));
      autoResizeColumns(sheet);
      return corsResponse({ ok: true });
    }

    // ── Update single FD ─────────────────────────────────────────────────
    if (action === 'update') {
      const rowIdx = findRowIndex(sheet, payload.fd.id);
      if (rowIdx === -1) return corsResponse({ ok: false, error: 'FD not found' });
      sheet.getRange(rowIdx, 1, 1, HEADERS.length).setValues([objectToRow(payload.fd)]);
      return corsResponse({ ok: true });
    }

    // ── Delete single FD ─────────────────────────────────────────────────
    if (action === 'delete') {
      const rowIdx = findRowIndex(sheet, payload.id);
      if (rowIdx === -1) return corsResponse({ ok: false, error: 'FD not found' });
      sheet.deleteRow(rowIdx);
      return corsResponse({ ok: true });
    }

    return corsResponse({ ok: false, error: 'Unknown action' });
  } catch(err) {
    return corsResponse({ ok: false, error: err.message });
  }
}

function autoResizeColumns(sheet) {
  sheet.autoResizeColumns(1, HEADERS.length);
}
