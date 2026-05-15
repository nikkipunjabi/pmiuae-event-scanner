/**
 * Event Check-In — Google Apps Script backend
 *
 * Deployed as a Web App, this script exposes two endpoints used by the
 * static QR scanner frontend:
 *
 *   GET  ?action=lookup&id=<registrant_id>
 *        → returns the registrant record as JSON
 *
 *   POST { action: "checkin", id: <registrant_id> }
 *        → sets "Checked In" = "Yes" and "Checked In Time" = now
 *          (in the spreadsheet's timezone), returns the updated record
 *
 * Setup:
 *   1. Open the sheet → Extensions → Apps Script
 *   2. Replace Code.gs with this file
 *   3. Update SHEET_ID and SHEET_NAME below if needed
 *   4. Deploy → New deployment → Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   5. Copy the resulting /exec URL into the frontend's config.js
 */

// ────────────────────────────────────────────────────────────────────
// CONFIG — change these if your sheet ID or tab name differs
// ────────────────────────────────────────────────────────────────────
const SHEET_ID   = '1Y7aRxKemiTLT1Llk-uq7p2X8b3eUxQ7yqJeSK3HGwBg';
const SHEET_NAME = 'Sheet1';

// Column letters in the registrants sheet (1-indexed numbers shown beside).
// These follow the header row you shared.
const COL = {
  ID:               1,  // A
  EVENT_ID:         2,  // B
  EVENT:            3,  // C
  EVENT_DATE:       4,  // D
  USER_ID:          5,  // E
  FIRST_NAME:       6,  // F
  LAST_NAME:        7,  // G
  EMAIL:            8,  // H
  NUM_REGISTRANTS:  9,  // I
  GROSS_AMOUNT:    10,  // J
  DISCOUNT_AMOUNT: 11,  // K
  LATE_FEE:        12,  // L
  NET_AMOUNT:      13,  // M
  COUPON:          14,  // N
  REG_DATE:        15,  // O
  TRANSACTION_ID:  16,  // P
  PAYMENT_DATE:    17,  // Q
  PAYMENT_STATUS:  18,  // R
  CHECKED_IN:      19,  // S
  CHECKED_IN_TIME: 20,  // T
};


// ════════════════════════════════════════════════════════════════════
// HTTP HANDLERS
// ════════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'lookup';

    if (action === 'lookup') {
      const id = e.parameter.id;
      if (!id) return jsonResponse({ ok: false, error: 'Missing id parameter' });
      const rec = findRegistrant(id);
      if (!rec) return jsonResponse({ ok: false, error: 'Registrant not found', id: id });
      return jsonResponse({ ok: true, registrant: rec });
    }

    // Bulk fetch — return every registrant for a given event so the
    // frontend can build a local lookup table and avoid round-trips.
    if (action === 'event') {
      const eventId = e.parameter.id || e.parameter.eventId;
      if (!eventId) return jsonResponse({ ok: false, error: 'Missing event id' });
      const registrants = getEventRegistrants(eventId);
      return jsonResponse({
        ok: true,
        eventId: eventId,
        count: registrants.length,
        registrants: registrants,
      });
    }

    if (action === 'ping') {
      // Pre-warm: touching the sheet here means the next real read is hot.
      try { getSheet_().getRange('A1').getValue(); } catch (_) {}
      return jsonResponse({ ok: true, message: 'pong', time: new Date().toISOString() });
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      body = e.parameter;
    }

    const action = body.action || 'checkin';
    const id = body.id;

    if (action === 'checkin') {
      if (!id) return jsonResponse({ ok: false, error: 'Missing id' });
      const result = checkInRegistrant(id, body.force === true);
      return jsonResponse(result);
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}


// ════════════════════════════════════════════════════════════════════
// CORE LOOKUP + UPDATE
// ════════════════════════════════════════════════════════════════════

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet tab not found: ' + SHEET_NAME);
  return sheet;
}

// Cache TTL: 5 minutes is a good balance for live events. The cache is
// invalidated on writes (check-in) so attendees who get checked in always
// reflect immediately on the device that did the check-in. Other devices
// might see slightly stale data for up to 5 min, but the server is the
// source of truth on the actual write.
const CACHE_TTL_SECONDS = 300;

function cache_() { return CacheService.getScriptCache(); }

/**
 * Read every data row in one batched call. This is the heaviest API
 * call we make, but doing it once is dramatically faster than doing
 * one read per scan. Memoised in CacheService for 5 min.
 */
function readAllRows_() {
  const c = cache_();
  const cached = c.get('all_rows_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, COL.CHECKED_IN_TIME).getValues();
  // Coerce dates to ISO strings so they survive JSON round-tripping.
  const cleaned = values.map(function (r) {
    return r.map(function (v) { return v instanceof Date ? v.toISOString() : v; });
  });

  try {
    // CacheService has a 100KB per-key limit; for large sheets we just
    // skip caching and fall back to the live read.
    c.put('all_rows_v1', JSON.stringify(cleaned), CACHE_TTL_SECONDS);
  } catch (_) {
    // payload too large — that's fine, we'll re-read next time
  }
  return cleaned;
}

function invalidateRowCache_() {
  try { cache_().remove('all_rows_v1'); } catch (_) {}
}

/**
 * Find a registrant by ID (column A). Returns { row, data: {…} } or null.
 * The match is string-based so numeric and text IDs both work.
 *
 * Performance: uses the in-memory row cache so repeat lookups don't
 * round-trip to SpreadsheetApp.
 */
function findRegistrant(id) {
  const all = readAllRows_();
  const target = String(id).trim();

  for (let i = 0; i < all.length; i++) {
    if (String(all[i][0]).trim() === target) {
      return {
        row: i + 2, // +2: header row + 0-indexed loop
        data: rowToObject_(all[i]),
      };
    }
  }
  return null;
}

/**
 * Bulk-fetch every registrant for a given event ID. Used by the frontend
 * to build a local lookup table on app load so subsequent scans don't
 * hit the network at all.
 */
function getEventRegistrants(eventId) {
  const all = readAllRows_();
  const target = String(eventId).trim();
  const out = [];

  for (let i = 0; i < all.length; i++) {
    if (String(all[i][COL.EVENT_ID - 1]).trim() === target) {
      out.push(rowToObject_(all[i]));
    }
  }
  return out;
}

function checkInRegistrant(id, force) {
  const sheet = getSheet_();
  const found = findRegistrant(id);
  if (!found) return { ok: false, error: 'Registrant not found', id: id };

  const existingCheckedIn = String(found.data.checkedIn || '').trim().toLowerCase();
  const alreadyCheckedIn = existingCheckedIn === 'yes' || existingCheckedIn === 'true';

  if (alreadyCheckedIn && !force) {
    return {
      ok: false,
      alreadyCheckedIn: true,
      previousTime: found.data.checkedInTime,
      registrant: found.data,
    };
  }

  const now = new Date();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'UTC';
  const formatted = Utilities.formatDate(now, tz, 'dd-MM-yyyy HH:mm:ss');

  // Single batched write — writes both cells in one Sheets API call.
  sheet.getRange(found.row, COL.CHECKED_IN, 1, 2).setValues([['Yes', formatted]]);

  // Invalidate the row cache so the next read sees our write.
  invalidateRowCache_();

  // Skip reading the row back — we already know what we wrote.
  const updated = Object.assign({}, found.data, {
    checkedIn: 'Yes',
    checkedInTime: formatted,
  });

  return {
    ok: true,
    overwritten: alreadyCheckedIn,
    registrant: updated,
  };
}


// ════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════

function rowToObject_(row) {
  return {
    id:              row[COL.ID - 1],
    eventId:         row[COL.EVENT_ID - 1],
    event:           row[COL.EVENT - 1],
    eventDate:       row[COL.EVENT_DATE - 1],
    userId:          row[COL.USER_ID - 1],
    firstName:       row[COL.FIRST_NAME - 1],
    lastName:        row[COL.LAST_NAME - 1],
    email:           row[COL.EMAIL - 1],
    numRegistrants:  row[COL.NUM_REGISTRANTS - 1],
    grossAmount:     row[COL.GROSS_AMOUNT - 1],
    discountAmount:  row[COL.DISCOUNT_AMOUNT - 1],
    lateFee:         row[COL.LATE_FEE - 1],
    netAmount:       row[COL.NET_AMOUNT - 1],
    coupon:          row[COL.COUPON - 1],
    registrationDate: row[COL.REG_DATE - 1],
    transactionId:   row[COL.TRANSACTION_ID - 1],
    paymentDate:     row[COL.PAYMENT_DATE - 1],
    paymentStatus:   row[COL.PAYMENT_STATUS - 1],
    checkedIn:       row[COL.CHECKED_IN - 1],
    checkedInTime:   row[COL.CHECKED_IN_TIME - 1],
  };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════════════
// OPTIONAL: run from the editor to test without deploying
// ════════════════════════════════════════════════════════════════════
function _testLookup() {
  Logger.log(findRegistrant('6845'));
}
function _testCheckIn() {
  Logger.log(checkInRegistrant('6845', false));
}


// ════════════════════════════════════════════════════════════════════
// EVENT SUMMARY TAB  — builds a reusable, formula-driven dashboard
// ════════════════════════════════════════════════════════════════════
//
// Usage:
//   1. From the Apps Script editor, run buildSummaryTab once (it will
//      ask for permission the first time).
//   2. Reload the spreadsheet — a new "Event Summary" menu appears in
//      the toolbar.
//   3. To pick a different event, just change the dropdown in cell B3
//      on the Summary tab — every metric recalculates automatically.
//   4. To rebuild from scratch (e.g. after schema changes), use
//      Event Summary → Build / Refresh.
//
// Default event when first built: 301 (AI-Driven PMOs…)
// ────────────────────────────────────────────────────────────────────

const SUMMARY_TAB_NAME   = 'Summary';
const SUMMARY_DEFAULT_ID = 301;

// ── colours ─────────────────────────────────────────────────────────
const C_TITLE_BG     = '#0f172a';
const C_TITLE_FG     = '#ffffff';
const C_SECTION_BG   = '#1e293b';
const C_SECTION_FG   = '#ffffff';
const C_PICKER_BG    = '#fef3c7';
const C_VALUE_GREEN  = '#15803d';
const C_VALUE_AMBER  = '#b45309';
const C_BORDER       = '#cbd5e1';
const C_HEADER_BG    = '#f1f5f9';


function onOpen() {
  SpreadsheetApp.getActive().addMenu('Event Summary', [
    { name: 'Build / Refresh Summary tab', functionName: 'buildSummaryTab' },
  ]);
}


function buildSummaryTab() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const source = ss.getSheetByName(SHEET_NAME);
  if (!source) throw new Error('Source tab not found: ' + SHEET_NAME);

  // Preserve the currently selected event ID if the tab already exists
  // (so a rebuild doesn't lose the user's selection).
  let preservedId = SUMMARY_DEFAULT_ID;
  const existing = ss.getSheetByName(SUMMARY_TAB_NAME);
  if (existing) {
    const v = existing.getRange('B3').getValue();
    if (v !== '' && v != null) preservedId = v;
  }

  // Wipe + recreate
  let s = existing;
  if (s) {
    s.clear();
    s.clearConditionalFormatRules();
    s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).clearDataValidations();
  } else {
    s = ss.insertSheet(SUMMARY_TAB_NAME);
  }
  s.setHiddenGridlines(true);

  // Column widths
  s.setColumnWidth(1, 180);
  s.setColumnWidth(2, 160);
  s.setColumnWidth(3, 110);
  s.setColumnWidth(4, 200);
  s.setColumnWidth(5, 240);
  s.setColumnWidth(6, 130);
  s.setColumnWidth(7, 130);
  s.setColumnWidth(8, 170);

  // ── Title row ────────────────────────────────────────────────────
  s.getRange('A1:H1').merge()
    .setValue('EVENT SUMMARY')
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor(C_TITLE_FG)
    .setBackground(C_TITLE_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  s.setRowHeight(1, 48);

  // ── Event picker ─────────────────────────────────────────────────
  s.getRange('A3').setValue('Event ID:').setFontWeight('bold').setFontSize(11);
  s.getRange('B3').setValue(preservedId)
    .setBackground(C_PICKER_BG)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#a16207', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  s.getRange('D3').setValue('Event Name:').setFontWeight('bold');
  s.getRange('E3').setFormula(
    '=IFERROR(INDEX(' + SHEET_NAME + '!C:C, MATCH(B3, ' + SHEET_NAME + '!B:B, 0)), "—")'
  ).setFontSize(11);
  s.getRange('E3:H3').merge().setWrap(true);

  s.getRange('D4').setValue('Event Date:').setFontWeight('bold');
  s.getRange('E4').setFormula(
    '=IFERROR(INDEX(' + SHEET_NAME + '!D:D, MATCH(B3, ' + SHEET_NAME + '!B:B, 0)), "—")'
  );

  // Build the dropdown using a hidden helper column (unique Event IDs).
  s.getRange('L1').setFormula(
    '=IFERROR(UNIQUE(FILTER(' + SHEET_NAME + '!B2:B, ' + SHEET_NAME + '!B2:B<>"")), "")'
  );
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(s.getRange('L1:L'), true)
    .setAllowInvalid(true) // allow typing an ID even if it's not yet in the sheet
    .setHelpText('Pick an event from the dropdown, or type any Event ID.')
    .build();
  s.getRange('B3').setDataValidation(rule);
  s.hideColumns(12); // hide column L

  // ── Sections ─────────────────────────────────────────────────────
  let row = 6;

  // HEADCOUNT  — capture each row in a named variable so formulas
  // can reference the right cells without relying on `row++` arithmetic
  // (which is bug-prone because the increment fires before the string
  // template is built).
  row = section_(s, row, 'HEADCOUNT');
  const rTotal       = row++;
  const rCheckedIn   = row++;
  const rNotCheckedIn = row++;
  const rCheckInRate = row++;
  metric_(s, rTotal,        'Total Registrants',   '=COUNTIF(' + SHEET_NAME + '!B:B, B3)');
  metric_(s, rCheckedIn,    'Checked In',          '=COUNTIFS(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!S:S, "Yes")', { color: C_VALUE_GREEN, bold: true });
  metric_(s, rNotCheckedIn, 'Not Yet Checked In',  '=B' + rTotal + ' - B' + rCheckedIn, { color: C_VALUE_AMBER });
  metric_(s, rCheckInRate,  'Check-In Rate',       '=IFERROR(B' + rCheckedIn + '/B' + rTotal + ', 0)', { numberFormat: '0.0%' });
  row++;

  // PAYMENT STATUS — same pattern. Percentages divide by the captured
  // Total Registrants row so this works no matter where HEADCOUNT lives.
  row = section_(s, row, 'PAYMENT STATUS');
  const rPaid    = row++;
  const rUnpaid  = row++;
  const rPending = row++;
  const rOther   = row++;
  metric_(s, rPaid,    'Paid',    '=COUNTIFS(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!R:R, "Paid")');
  s.getRange('C' + rPaid).setFormula('=IFERROR(B' + rPaid + '/B' + rTotal + ', 0)').setNumberFormat('0.0%').setFontColor('#475569');
  metric_(s, rUnpaid,  'Unpaid',  '=COUNTIFS(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!R:R, "Unpaid")');
  s.getRange('C' + rUnpaid).setFormula('=IFERROR(B' + rUnpaid + '/B' + rTotal + ', 0)').setNumberFormat('0.0%').setFontColor('#475569');
  metric_(s, rPending, 'Pending', '=COUNTIFS(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!R:R, "Pending")');
  s.getRange('C' + rPending).setFormula('=IFERROR(B' + rPending + '/B' + rTotal + ', 0)').setNumberFormat('0.0%').setFontColor('#475569');
  metric_(s, rOther,   'Other',   '=B' + rTotal + ' - SUM(B' + rPaid + ':B' + rPending + ')');
  row++;

  // REVENUE
  row = section_(s, row, 'REVENUE');
  metric_(s, row++, 'Gross Amount',    '=SUMIF(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!J:J)', { numberFormat: '#,##0.00' });
  metric_(s, row++, 'Discount Amount', '=SUMIF(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!K:K)', { numberFormat: '#,##0.00' });
  metric_(s, row++, 'Late Fee',        '=SUMIF(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!L:L)', { numberFormat: '#,##0.00' });
  metric_(s, row++, 'Net Revenue',     '=SUMIF(' + SHEET_NAME + '!B:B, B3, ' + SHEET_NAME + '!M:M)', { numberFormat: '#,##0.00', bold: true });
  row++;

  // ── ATTENDEE LIST ────────────────────────────────────────────────
  row = section_(s, row, 'ATTENDEE LIST');
  const Q = "'";  // single-quote helper for QUERY label syntax
  const attendeeFormula =
    '=IFERROR(' +
      'QUERY(' + SHEET_NAME + '!A2:T, ' +
        '"select A, F, G, H, O, R, S, T ' +
         'where B = " & B3 & " ' +
         'order by A desc ' +
         'label A ' + Q + 'ID' + Q + ', ' +
                'F ' + Q + 'First Name' + Q + ', ' +
                'G ' + Q + 'Last Name'  + Q + ', ' +
                'H ' + Q + 'Email'      + Q + ', ' +
                'O ' + Q + 'Registration Date' + Q + ', ' +
                'R ' + Q + 'Payment Status'    + Q + ', ' +
                'S ' + Q + 'Checked In'        + Q + ', ' +
                'T ' + Q + 'Check-In Time'     + Q + '"), ' +
    '"No registrants for this event yet.")';
  s.getRange('A' + row).setFormula(attendeeFormula);

  // Style the header row that QUERY spills out.
  const headerRow = row;
  s.getRange(headerRow, 1, 1, 8)
    .setFontWeight('bold')
    .setBackground(C_HEADER_BG)
    .setBorder(true, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // Conditional formatting on the attendee list
  const tableRange   = s.getRange(headerRow + 1, 7, 1000, 1); // column G = "Checked In"
  const paymentRange = s.getRange(headerRow + 1, 6, 1000, 1); // column F = "Payment Status"
  const rules = s.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Yes').setBackground('#dcfce7').setFontColor('#14532d')
      .setRanges([tableRange]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('No').setBackground('#fee2e2').setFontColor('#7f1d1d')
      .setRanges([tableRange]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Paid').setBackground('#dcfce7').setFontColor('#14532d').setBold(true)
      .setRanges([paymentRange]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Unpaid').setBackground('#fee2e2').setFontColor('#7f1d1d').setBold(true)
      .setRanges([paymentRange]).build()
  );
  s.setConditionalFormatRules(rules);

  // Freeze the area above the attendee list so it stays visible while scrolling.
  s.setFrozenRows(headerRow);

  // Tidy alignment on the metric block
  s.getRange('A6:A' + (headerRow - 2)).setVerticalAlignment('middle');
  s.getRange('B6:B' + (headerRow - 2)).setHorizontalAlignment('right').setVerticalAlignment('middle');

  // Activate the tab and select the picker
  s.activate();
  s.getRange('B3').activate();

  SpreadsheetApp.getActive().toast(
    'Summary tab built. Change cell B3 to switch events.',
    'Event Summary',
    5
  );
}


// ── helpers ─────────────────────────────────────────────────────────
function section_(sheet, row, label) {
  sheet.getRange(row, 1, 1, 8).merge()
    .setValue(label)
    .setFontWeight('bold')
    .setFontColor(C_SECTION_FG)
    .setBackground(C_SECTION_BG)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setFontSize(11);
  sheet.setRowHeight(row, 28);
  return row + 1;
}

function metric_(sheet, row, label, formula, opts) {
  opts = opts || {};
  sheet.getRange(row, 1).setValue(label).setFontWeight('normal');
  const valueCell = sheet.getRange(row, 2);
  valueCell.setFormula(formula).setHorizontalAlignment('right');
  if (opts.color)        valueCell.setFontColor(opts.color);
  if (opts.bold)         valueCell.setFontWeight('bold');
  if (opts.numberFormat) valueCell.setNumberFormat(opts.numberFormat);
  sheet.setRowHeight(row, 26);
}
