// ============================================================
// Cash Drawer Manager - Google Apps Script
// ============================================================
// SHEET SETUP:
//   Tab 1: "Drawer"  - denomination inventory + thresholds
//   Tab 2: "Log"     - transaction history
//
// Drawer tab layout (row 1 = headers, row 2+ = data):
//   Left table (current cash on hand)
//     A: Denomination   B: Count   C: Value (A×B)
//   D: (blank spacer)
//   Right table (thresholds / targets)
//     E: Denomination   F: Min Warning   G: Min Critical   H: Required Value (E×F)
//
// Log tab columns:
//   A: Timestamp  B: Action  C: 20  D: 10  E: 5  F: 1  G: Total Change  H: Note  I: Status
// ============================================================

const DRAWER_SHEET = 'Drawer';
const LOG_SHEET = 'Log';
const DENOMINATIONS = [100, 50, 20, 10, 5, 2, 1]; // order must match Log tab columns C-F

const API_KEY_USERS = {
    'Levi-APIKEY': 'Levi',
    'Kate-APIKEY': 'Kate',
    'Noah-APIKEY': 'Noah',
  };

// ---- Utility -----------------------------------------------

function getDrawerSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DRAWER_SHEET);
}

function getLogSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
}

/**
 * Reads the Drawer tab into a map keyed by denomination.
 * Returns: { 20: { row, count, minWarning, minCritical }, ... }
 */
function readDrawer() {
  const sheet = getDrawerSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return {};
  }
  // Read columns A–H so we can get counts and threshold values
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const map = {};
  data.forEach((row, i) => {
    const denom = Number(row[0]);            // left table denomination (col A)
    if (DENOMINATIONS.includes(denom)) {
      map[denom] = {
        row: i + 2, // 1-indexed sheet row
        count: Number(row[1]),        // Count column (B)
        minWarning: Number(row[5]),   // Min Warning column (F)
        minCritical: Number(row[6]),  // Min Critical column (G)
      };
    }
  });
  return map;
}

/**
 * Evaluates status across all denominations.
 * Returns: { status, warnings: [{ denom, count, level }] }
 */
function evaluateStatus(drawer) {
  let status = 'success';
  const warnings = [];

  DENOMINATIONS.forEach(denom => {
    const d = drawer[denom];
    if (!d) return;

    if (d.minCritical > 0 && d.count <= d.minCritical) {
      warnings.push({ denom, count: d.count, level: 'critical' });
      status = 'critical';
    } else if (d.minWarning > 0 && d.count <= d.minWarning) {
      warnings.push({ denom, count: d.count, level: 'low' });
      if (status !== 'critical') status = 'low';
    }
  });

  return { status, warnings };
}

/**
 * Calculates total cash value of a drawer map.
 */
function totalValue(drawer) {
  return DENOMINATIONS.reduce((sum, denom) => {
    return sum + (drawer[denom] ? drawer[denom].count * denom : 0);
  }, 0);
}

/**
 * Calculates the minimum total required based on minWarning thresholds.
 * Only counts denominations where minWarning > 0.
 */
function minimumTotal(drawer) {
  return DENOMINATIONS.reduce((sum, denom) => {
    const d = drawer[denom];
    return sum + (d && d.minWarning > 0 ? d.minWarning * denom : 0);
  }, 0);
}

/**
 * Appends a row to the Log tab.
 */
function appendLog(action, bills, totalChange, note, status) {
  const sheet = getLogSheet();
  const row = [
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss'),
    action,
    bills[100] || 0,
    bills[50] || 0,
    bills[20] || 0,
    bills[10] || 0,
    bills[5]  || 0,
    bills[2]  || 0,
    bills[1]  || 0,
    totalChange,
    note || '',
    status,
  ];
  sheet.appendRow(row);
}

// ---- Request Handlers --------------------------------------

/**
 * Handles delta transactions: adds/subtracts bill counts.
 */
function handleDelta(bills, note) {
  // Reject if all deltas are zero — nothing would change
  const hasChange = Object.values(bills).some(v => Number(v) !== 0);
  if (!hasChange) {
    return errorResponse('No changes — all bill counts are 0');
  }

  const sheet = getDrawerSheet();
  const drawer = readDrawer();

  // Validate — don't allow counts to go negative
  for (const [denomStr, delta] of Object.entries(bills)) {
    const denom = Number(denomStr);
    if (!drawer[denom]) {
      return errorResponse(`Unknown denomination: ${denomStr}`);
    }
    const newCount = drawer[denom].count + Number(delta);
    if (newCount < 0) {
      return errorResponse(`Cannot remove ${Math.abs(delta)} $${denom} bills — only ${drawer[denom].count} in drawer`);
    }
  }

  // Apply deltas
  let totalChange = 0;
  for (const [denomStr, delta] of Object.entries(bills)) {
    const denom = Number(denomStr);
    const d = drawer[denom];
    const newCount = d.count + Number(delta);
    sheet.getRange(d.row, 2).setValue(newCount);
    d.count = newCount; // update local map for status check
    totalChange += Number(delta) * denom;
  }

  const { status, warnings } = evaluateStatus(drawer);
  const drawerTotal = totalValue(drawer);

  appendLog('delta', bills, totalChange, note, status);

  return buildResponse(status, drawer, drawerTotal, warnings, totalChange);
}

/**
 * Handles absolute set: replaces bill counts entirely.
 */
function handleSet(bills, note) {
  const sheet = getDrawerSheet();
  const drawer = readDrawer();

  for (const [denomStr] of Object.entries(bills)) {
    if (!drawer[Number(denomStr)]) {
      return errorResponse(`Unknown denomination: ${denomStr}`);
    }
  }

  let totalChange = 0;
  for (const [denomStr, newCount] of Object.entries(bills)) {
    const denom = Number(denomStr);
    const d = drawer[denom];
    if (Number(newCount) < 0) {
      return errorResponse(`Count cannot be negative for $${denom}`);
    }
    const delta = Number(newCount) - d.count;
    sheet.getRange(d.row, 2).setValue(Number(newCount));
    totalChange += delta * denom;
    d.count = Number(newCount);
  }

  const { status, warnings } = evaluateStatus(drawer);
  const drawerTotal = totalValue(drawer);

  appendLog('set', bills, totalChange, note, status);

  return buildResponse(status, drawer, drawerTotal, warnings, totalChange);
}

// ---- Response Builders -------------------------------------

function buildResponse(status, drawer, drawerTotal, warnings, totalChange) {
  // Map status to a short human label
  const label = status === 'success' ? 'good'
              : status === 'low'     ? 'warning'
              : 'bad';

  const payload = {
    result: 'success',
    cash: label,                     // "good" | "warning" | "bad"
    total: `$${drawerTotal}`,
    minTotal: `$${minimumTotal(drawer)}`,
  };

  if (totalChange !== 0) {
    payload.change = `$${totalChange}`;
  }

  if (warnings.length > 0) {
    payload.alerts = warnings.map(w =>
      w.level === 'critical'
        ? `$${w.denom} OUT (${w.count})`
        : `$${w.denom} LOW (${w.count})`
    );
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'error', message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- GET: Status Check -------------------------------------

function doGet(e) {
  try {
    // SECURITY CHECK: Verify API key (expects ?apiKey=... in query params)
    const apiKey = e && e.parameter && e.parameter.apiKey;
    if (!apiKey || !API_KEY_USERS[apiKey]) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: 'Invalid or missing API key.'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    const drawer = readDrawer();
    const { status, warnings } = evaluateStatus(drawer);
    const drawerTotal = totalValue(drawer);
    return buildResponse(status, drawer, drawerTotal, warnings, 0);
  } catch (err) {
    return errorResponse(err.message);
  }
}

// ---- POST: Transaction -------------------------------------

/**
 * Expected POST body (JSON):
 * {
 *   "action": "delta" | "set",   // optional, defaults to "delta"
 *   "note": "optional note",
 *   "bills": {
 *     "20": 2,    // positive = added, negative = removed (delta)
 *     "5": -3     // omit denominations that didn't change
 *   }
 * }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    // SECURITY CHECK: Verify API key (expects {"apiKey": "..."} in JSON body)
    const apiKey = body.apiKey;
    if (!apiKey || !API_KEY_USERS[apiKey]) {
      return ContentService.createTextOutput(JSON.stringify({
        result: 'error',
        message: 'Invalid or missing API key.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const action = (body.action || 'delta').toLowerCase();
    const bills = body.bills;
    const note = body.note || '';

    if (!bills || typeof bills !== 'object' || Object.keys(bills).length === 0) {
      return errorResponse('bills object is required and cannot be empty');
    }

    // VALIDATION: ensure only real bill denominations and integer counts
    const validationError = validateBills(bills, action);
    if (validationError) {
      return errorResponse(validationError);
    }

    if (action === 'set') {
      return handleSet(bills, note);
    } else if (action === 'delta') {
      return handleDelta(bills, note);
    } else {
      return errorResponse(`Unknown action: "${action}". Use "delta" or "set".`);
    }

  } catch (err) {
    return errorResponse(`Failed to parse request: ${err.message}`);
  }
}

/**
 * Validates that the incoming bills object only contains real bill
 * denominations and integer counts.
 *
 * @param {Object} bills - e.g. { "20": 1, "5": -3 }
 * @param {string} action - "delta" | "set"
 * @return {string|null} error message if invalid, otherwise null
 */
function validateBills(bills, action) {
  const validDenoms = new Set(DENOMINATIONS.map(String));

  for (const [denomStr, value] of Object.entries(bills)) {
    // Denomination must be one of the supported bills
    if (!validDenoms.has(denomStr)) {
      return `Invalid bill denomination: $${denomStr}. Allowed bills: ${DENOMINATIONS.join(', ')}`;
    }

    const num = Number(value);
    if (!isFinite(num)) {
      return `Bill count for $${denomStr} must be a number. Received: ${value}`;
    }

    if (!Number.isInteger(num)) {
      return `Bill count for $${denomStr} must be a whole number. Received: ${value}`;
    }

    if (action === 'set' && num < 0) {
      return `Bill count for $${denomStr} cannot be negative when using "set". Received: ${value}`;
    }
  }

  return null;
}

// ---- Setup Helper ------------------------------------------

/**
 * Run this once manually to create/format the Drawer and Log tabs.
 * Safe to re-run — won't overwrite existing data rows.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Drawer tab ---
  let drawer = ss.getSheetByName(DRAWER_SHEET);
  if (!drawer) drawer = ss.insertSheet(DRAWER_SHEET);

  // Headers: split into two visual tables (left: live counts, right: thresholds)
  const drawerHeaderRange = drawer.getRange(1, 1, 1, 8);
  drawerHeaderRange
    .setValues([[
      'Denomination', 'Count', 'Value', '', 'Denomination', 'Min Warning', 'Min Critical', 'Required Value'
    ]])
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  // Left table header: dark blue
  drawer.getRange(1, 1, 1, 3).setBackground('#1A56DB');
  // Right table header: dark slate
  drawer.getRange(1, 5, 1, 4).setBackground('#374151');
  drawerHeaderRange.setBorder(true, true, true, true, true, true);
  // Freeze header row so it stays visible above data
  drawer.setFrozenRows(1);

  // Default data rows if sheet is empty
  if (drawer.getLastRow() < 2) {
    // [denomination, count, minWarning, minCritical]
    const defaults = [
      [100, 0,  0, 0],
      [50,  0,  0, 0],
      [20, 10, 5, 2],
      [10, 10, 4, 1],
      [5,  20, 8, 2],
      [2,  0,  0, 0],
      [1,  30, 10, 3],
    ];

    // Left table: denomination + count (cols A,B)
    drawer.getRange(2, 1, defaults.length, 2)
      .setValues(defaults.map(r => [r[0], r[1]]));

    // Right table: denomination + thresholds (cols E,F,G)
    drawer.getRange(2, 5, defaults.length, 3)
      .setValues(defaults.map(r => [r[0], r[2], r[3]]));
  }

  // Data rows for formulas (only over actual denominations, not totals)
  const firstDataRow = 2;
  const lastDataRow = drawer.getLastRow();
  const numDataRows = lastDataRow - firstDataRow + 1;

  // Left table "Value" column: =A(row) * B(row)
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    drawer.getRange(r, 3).setFormula(`=A${r}*B${r}`);
  }

  // Right table "Required Value" column: =E(row) * F(row)
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    drawer.getRange(r, 8).setFormula(`=E${r}*F${r}`);
  }

  // Dollar formatting: Denomination cols (A, E) and Value cols (C, H)
  const dollarFmt = '"$"#,##0';
  drawer.getRange(firstDataRow, 1, numDataRows, 1).setNumberFormat(dollarFmt); // col A: denomination
  drawer.getRange(firstDataRow, 3, numDataRows, 1).setNumberFormat(dollarFmt); // col C: value
  drawer.getRange(firstDataRow, 5, numDataRows, 1).setNumberFormat(dollarFmt); // col E: denomination
  drawer.getRange(firstDataRow, 8, numDataRows, 1).setNumberFormat(dollarFmt); // col H: required value

  // Alternating row backgrounds for left table (cols A–C)
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const bg = (r % 2 === 0) ? '#EFF6FF' : '#FFFFFF'; // light blue stripe / white
    drawer.getRange(r, 1, 1, 3).setBackground(bg);
  }

  // Alternating row backgrounds for right table (cols E–H)
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const bg = (r % 2 === 0) ? '#F3F4F6' : '#FFFFFF'; // light grey stripe / white
    drawer.getRange(r, 5, 1, 4).setBackground(bg);
  }

  // Outer border around left table (rows 1–lastData, cols A–C)
  drawer.getRange(1, 1, lastDataRow, 3)
    .setBorder(true, true, true, true, null, null,
      '#1A56DB', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Outer border around right table (rows 1–lastData, cols E–H)
  drawer.getRange(1, 5, lastDataRow, 4)
    .setBorder(true, true, true, true, null, null,
      '#374151', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Inner grid lines for both tables
  drawer.getRange(firstDataRow, 1, numDataRows, 3)
    .setBorder(null, null, null, null, true, true,
      '#BFDBFE', SpreadsheetApp.BorderStyle.SOLID);
  drawer.getRange(firstDataRow, 5, numDataRows, 4)
    .setBorder(null, null, null, null, true, true,
      '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);

  // Total row — always row 9 (fixed after 7 denomination rows, 2–8)
  // Clear any stale total rows beyond the data before writing
  const totalRow = lastDataRow + 1;
  drawer.getRange(totalRow, 1, drawer.getMaxRows() - lastDataRow, 3).clearContent().clearFormat();
  drawer.getRange(totalRow, 1, 1, 3)
    .setBackground('#1A56DB')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, null, null,
      '#1A56DB', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  drawer.getRange(totalRow, 1).setValue('TOTAL');
  drawer.getRange(totalRow, 2).setFormula(`=SUM(B${firstDataRow}:B${lastDataRow})`);
  drawer.getRange(totalRow, 3).setFormula(`=SUM(C${firstDataRow}:C${lastDataRow})`)
    .setNumberFormat('"$"#,##0');

  // Conditional formatting on Count column (col B): green / orange / red
  // Clear existing rules first
  drawer.clearConditionalFormatRules();
  const countRange = drawer.getRange(firstDataRow, 2, numDataRows, 1);
  const rules = [];

  // Critical: count <= minCritical (col G) — red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(G${firstDataRow}>0, B${firstDataRow}<=G${firstDataRow})`)
    .setBackground('#FEE2E2')
    .setFontColor('#991B1B')
    .setRanges([countRange])
    .build());

  // Warning: count <= minWarning (col F) — orange
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(F${firstDataRow}>0, B${firstDataRow}<=F${firstDataRow})`)
    .setBackground('#FEF3C7')
    .setFontColor('#92400E')
    .setRanges([countRange])
    .build());

  // Good: count > minWarning (or minWarning is 0) — green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR(F${firstDataRow}=0, B${firstDataRow}>F${firstDataRow})`)
    .setBackground('#D1FAE5')
    .setFontColor('#065F46')
    .setRanges([countRange])
    .build());

  drawer.setConditionalFormatRules(rules);

  // --- Log tab ---
  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) log = ss.insertSheet(LOG_SHEET);

  const logHeaderRange = log.getRange(1, 1, 1, 12);
  logHeaderRange
    .setValues([[
      'Timestamp', 'Action', '$100', '$50', '$20', '$10', '$5', '$2', '$1', 'Total Change ($)', 'Note', 'Status'
    ]])
    .setFontWeight('bold')
    .setBackground('#E8F0FE')
    .setFontColor('#000000')
    .setHorizontalAlignment('center');
  logHeaderRange.setBorder(true, true, true, true, true, true);
  log.setFrozenRows(1);

  // Center Timestamp (col A) and Action (col B) data cells
  log.getRange(2, 1, log.getMaxRows() - 1, 2).setHorizontalAlignment('center');

  SpreadsheetApp.getUi().alert('Setup complete! Drawer and Log tabs are ready.');
}

// ---- Debug Menu --------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cash Drawer')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Check Status', 'debugStatus')
    .addToUi();
}

function debugStatus() {
  const drawer = readDrawer();
  const { status, warnings } = evaluateStatus(drawer);
  const total = totalValue(drawer);
  const msg = [
    `Status: ${status.toUpperCase()}`,
    `Total Cash: $${total}`,
    warnings.length ? '\nWarnings:\n' + warnings.map(w => `  $${w.denom}: ${w.level} (${w.count} left)`).join('\n') : 'No warnings.',
  ].join('\n');
  SpreadsheetApp.getUi().alert(msg);
}