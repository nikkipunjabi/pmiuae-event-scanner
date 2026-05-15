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

    if (action === 'ping') {
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

/**
 * Find a registrant by ID (column A). Returns { row, data: {…} } or null.
 * The match is string-based so numeric and text IDs both work.
 */
function findRegistrant(id) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  // Read just column A to find the row, then fetch that row's data.
  const idValues = sheet.getRange(2, COL.ID, lastRow - 1, 1).getValues();
  const target = String(id).trim();

  let rowIndex = -1;
  for (let i = 0; i < idValues.length; i++) {
    if (String(idValues[i][0]).trim() === target) {
      rowIndex = i + 2; // +2 because header row + 0-based loop
      break;
    }
  }
  if (rowIndex === -1) return null;

  const row = sheet.getRange(rowIndex, 1, 1, COL.CHECKED_IN_TIME).getValues()[0];
  return {
    row: rowIndex,
    data: rowToObject_(row),
  };
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

  sheet.getRange(found.row, COL.CHECKED_IN).setValue('Yes');
  sheet.getRange(found.row, COL.CHECKED_IN_TIME).setValue(formatted);

  // Read the row back so the client sees the canonical values.
  const updatedRow = sheet.getRange(found.row, 1, 1, COL.CHECKED_IN_TIME).getValues()[0];
  return {
    ok: true,
    overwritten: alreadyCheckedIn,
    registrant: rowToObject_(updatedRow),
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
