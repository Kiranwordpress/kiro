// ════════════════════════════════════════════════════════════════════════════
//  FD Tracker — Google Apps Script Backend
//  Paste this entire file into Extensions → Apps Script in your Google Sheet
//  then Deploy → Manage Deployments → create a NEW deployment (do not reuse
//  the old one, otherwise changes won't take effect).
//  Set: Execute as = Me, Who has access = Anyone
// ════════════════════════════════════════════════════════════════════════════

const SHEET_NAME = 'FD_Data';

// Raw input columns (stored as-is)
const INPUT_HEADERS = [
  'id', 'bank', 'fdNum', 'principal', 'rate',
  'startDate', 'maturityDate', 'compounding', 'notes',
  'renewedFromId', 'renewedToId', 'createdAt'
];

// Calculated columns (computed server-side on every save)
const CALC_HEADERS = [
  'daysElapsed', 'totalDays', 'progress_pct',
  'currentValue', 'interestEarned', 'maturityValue', 'status'
];

const ALL_HEADERS = [...INPUT_HEADERS, ...CALC_HEADERS];

// ── Calculation helpers ───────────────────────────────────────────────────────
function calcFD(fd) {
  const P          = parseFloat(fd.principal) || 0;
  const r          = (parseFloat(fd.rate) || 0) / 100;
  const n          = parseInt(fd.compounding) || 4;
  const start      = new Date(fd.startDate);
  const maturity   = new Date(fd.maturityDate);
  const now        = new Date();

  if (isNaN(start) || isNaN(maturity) || P <= 0) {
    return { daysElapsed: 0, totalDays: 0, progress_pct: 0,
             currentValue: P, interestEarned: 0, maturityValue: P, status: 'Unknown' };
  }

  const totalDays   = Math.max((maturity - start) / 86400000, 1);
  const totalYears  = totalDays / 365;
  const maturityVal = P * Math.pow(1 + r / n, n * totalYears);

  const effectiveNow  = now > maturity ? maturity : now < start ? start : now;
  const daysElapsed   = Math.max(Math.floor((effectiveNow - start) / 86400000), 0);
  const yearsElapsed  = daysElapsed / 365;
  const currentVal    = P * Math.pow(1 + r / n, n * yearsElapsed);
  const progress      = Math.min(100, Math.round((daysElapsed / totalDays) * 100));

  // Status
  let status = 'Active';
  if (fd.renewedToId && fd.renewedToId !== '' && fd.renewedToId !== 'null') {
    status = 'Renewed';
  } else if (now >= maturity) {
    status = 'Matured';
  } else if (Math.ceil((maturity - now) / 86400000) <= 30) {
    status = 'Upcoming';
  }

  return {
    daysElapsed:   Math.round(daysElapsed),
    totalDays:     Math.round(totalDays),
    progress_pct:  progress,
    currentValue:  Math.round(currentVal * 100) / 100,
    interestEarned: Math.round((currentVal - P) * 100) / 100,
    maturityValue:  Math.round(maturityVal * 100) / 100,
    status,
  };
}

// ── Sheet helpers ─────────────────────────────────────────────────────────────
function getSheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    writeHeader(sheet);
  }
  return sheet;
}

function writeHeader(sheet) {
  sheet.getRange(1, 1, 1, ALL_HEADERS.length).setValues([ALL_HEADERS]);

  // Style input columns (blue)
  sheet.getRange(1, 1, 1, INPUT_HEADERS.length)
       .setFontWeight('bold')
       .setBackground('#2563eb')
       .setFontColor('#ffffff');

  // Style calculated columns (green)
  sheet.getRange(1, INPUT_HEADERS.length + 1, 1, CALC_HEADERS.length)
       .setFontWeight('bold')
       .setBackground('#16a34a')
       .setFontColor('#ffffff');

  sheet.setFrozenRows(1);
}

function ensureHeader(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, ALL_HEADERS.length).getValues()[0];
  // Re-write header if columns are missing or outdated
  if (firstRow.length < ALL_HEADERS.length || firstRow[0] !== 'id') {
    writeHeader(sheet);
  }
}

function fdToRow(fd) {
  const calc = calcFD(fd);
  const inputPart = INPUT_HEADERS.map(h => fd[h] !== undefined ? fd[h] : '');
  const calcPart  = CALC_HEADERS.map(h  => calc[h] !== undefined ? calc[h] : '');
  return [...inputPart, ...calcPart];
}

function rowToObject(row) {
  const obj = {};
  ALL_HEADERS.forEach((h, i) => obj[h] = row[i] === undefined ? '' : String(row[i]));
  return obj;
}

function getAllRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(rowToObject);
}

function findRowIndex(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function autoResizeColumns(sheet) {
  sheet.autoResizeColumns(1, ALL_HEADERS.length);
}

// ── Formatting helpers for sheet display ─────────────────────────────────────
function applyRowFormatting(sheet, startRow, numRows) {
  if (numRows <= 0) return;

  // Currency formatting for value columns
  const currencyCols = [
    ALL_HEADERS.indexOf('principal') + 1,
    ALL_HEADERS.indexOf('currentValue') + 1,
    ALL_HEADERS.indexOf('interestEarned') + 1,
    ALL_HEADERS.indexOf('maturityValue') + 1,
  ];
  currencyCols.forEach(col => {
    sheet.getRange(startRow, col, numRows, 1)
         .setNumberFormat('₹#,##,##0.00');
  });

  // Percentage formatting
  const pctCol = ALL_HEADERS.indexOf('rate') + 1;
  sheet.getRange(startRow, pctCol, numRows, 1)
       .setNumberFormat('0.00"%"');

  // Progress % column
  const progCol = ALL_HEADERS.indexOf('progress_pct') + 1;
  sheet.getRange(startRow, progCol, numRows, 1)
       .setNumberFormat('0"%"');

  // Alternating row colors
  for (let r = 0; r < numRows; r++) {
    const rowNum = startRow + r;
    const bg     = r % 2 === 0 ? '#f8fafc' : '#ffffff';
    sheet.getRange(rowNum, 1, 1, ALL_HEADERS.length).setBackground(bg);
  }
}

// ── Routes ────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'getAll';
    const sheet  = getSheet();
    ensureHeader(sheet);

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
    ensureHeader(sheet);

    // ── Save all (bulk sync) ────────────────────────────────────────────
    if (action === 'saveAll') {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

      if (payload.data && payload.data.length > 0) {
        const rows = payload.data.map(fdToRow);
        sheet.getRange(2, 1, rows.length, ALL_HEADERS.length).setValues(rows);
        applyRowFormatting(sheet, 2, rows.length);
      }
      autoResizeColumns(sheet);
      return corsResponse({ ok: true, count: (payload.data || []).length });
    }

    // ── Add single FD ──────────────────────────────────────────────────
    if (action === 'add') {
      const newRow = sheet.getLastRow() + 1;
      sheet.appendRow(fdToRow(payload.fd));
      applyRowFormatting(sheet, newRow, 1);
      autoResizeColumns(sheet);
      return corsResponse({ ok: true });
    }

    // ── Update single FD ───────────────────────────────────────────────
    if (action === 'update') {
      const rowIdx = findRowIndex(sheet, payload.fd.id);
      if (rowIdx === -1) return corsResponse({ ok: false, error: 'FD not found' });
      sheet.getRange(rowIdx, 1, 1, ALL_HEADERS.length).setValues([fdToRow(payload.fd)]);
      applyRowFormatting(sheet, rowIdx, 1);
      return corsResponse({ ok: true });
    }

    // ── Delete single FD ───────────────────────────────────────────────
    if (action === 'delete') {
      const rowIdx = findRowIndex(sheet, payload.id);
      if (rowIdx === -1) return corsResponse({ ok: false, error: 'FD not found' });
      sheet.deleteRow(rowIdx);
      return corsResponse({ ok: true });
    }

    // ── Recalculate all rows (refresh calculated columns) ──────────────
    if (action === 'recalculate') {
      const data = getAllRows(sheet);
      if (data.length === 0) return corsResponse({ ok: true, count: 0 });
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      const rows = data.map(fdToRow);
      sheet.getRange(2, 1, rows.length, ALL_HEADERS.length).setValues(rows);
      applyRowFormatting(sheet, 2, rows.length);
      autoResizeColumns(sheet);
      return corsResponse({ ok: true, count: data.length });
    }

    return corsResponse({ ok: false, error: 'Unknown action' });
  } catch(err) {
    return corsResponse({ ok: false, error: err.message });
  }
}
