/**
 * ============================================================
 *  SPF OF FINANCE DIRECTORATE SYSTEM — Google Apps Script
 *  Backend for Google Sheets sync (Officers + Transfers)
 * ============================================================
 *  Sida loo isticmaalo:
 *  1. Fur Google Sheet cusub (ama mid hore u haysatay)
 *  2. Extensions ➜ Apps Script
 *  3. Tirtir koodhkii hore ee Code.gs, ku dheji koodhkan oo dhan
 *  4. Deploy ➜ New deployment ➜ Type: Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  5. Riix Deploy, kadibna koobi URL-ka "Web app" ee soo baxa
 *  6. URL-kaas ku dheji System-ka SPF: Settings ➜ Google Sheets
 * ============================================================
 */

// ── CONFIG ──────────────────────────────────────────────────
var OFFICERS_SHEET_NAME  = 'Officers';
var TRANSFERS_SHEET_NAME = 'Transfers';

// Column order saved to the "Officers" sheet (matches SPF system record shape)
var OFFICER_COLUMNS = [
  'spfNum','name','apno','plno','phone','sex','rank','title',
  'station','qeyb','region','district','kash','dob','pob','tribe',
  'mother','marital','age','clan','bank','accno','salary','allow',
  'deduct','net','tdate','status','approvalStatus','approvalCreatedAt',
  'createdAt','updatedAt'
];

// Column order for the "Transfers" sheet
var TRANSFER_COLUMNS = [
  'id','officerSpf','officerName','officerApno','fromStation','toStation',
  'requestedBy','status','autoApproved','reason','createdAt','updatedAt'
];


// ── ENTRY POINTS ────────────────────────────────────────────

/**
 * GET — soo celisaa dhammaan xogta (officers + transfers) sida JSON.
 * Tijaabo: fur URL-ka browser-ka, waa inuu ku soo celiyo {"success":true,...}
 */
function doGet(e) {
  try {
    var type = (e && e.parameter && e.parameter.type) || 'all';
    var result = { success: true, updatedAt: new Date().toISOString() };

    if (type === 'officers' || type === 'all') {
      result.officers = readSheetAsObjects_(OFFICERS_SHEET_NAME, OFFICER_COLUMNS);
    }
    if (type === 'transfers' || type === 'all') {
      result.transfers = readSheetAsObjects_(TRANSFERS_SHEET_NAME, TRANSFER_COLUMNS);
    }
    // Backwards-compatible "data" field for the system's Test & Connect check
    result.data = result.officers || result.transfers || {};

    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}

/**
 * POST — qaata array xog ah (officers ama transfers), kuna qoraa/cusboonaysiiyaa
 * Sheet-ka. Haddii spfNum/id horeyba jiro, waa la cusboonaysiiyaa (update),
 * haddii kalena waa lagu daraa saf cusub (append).
 *
 * Body format laga filayo system-ka SPF:
 *   POST <url>                -> array of officer records (legacy/simple mode)
 *   POST <url>?type=officers  -> array of officer records
 *   POST <url>?type=transfers -> array of transfer records
 */
function doPost(e) {
  try {
    var type = (e && e.parameter && e.parameter.type) || 'officers';
    var body = JSON.parse(e.postData.contents || '[]');
    var records = Array.isArray(body) ? body : [body];

    var count = 0;
    if (type === 'transfers') {
      count = upsertRecords_(TRANSFERS_SHEET_NAME, TRANSFER_COLUMNS, records, 'id');
    } else {
      count = upsertRecords_(OFFICERS_SHEET_NAME, OFFICER_COLUMNS, records, 'spfNum');
    }

    return jsonOut_({ success: true, written: count, updatedAt: new Date().toISOString() });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}


// ── HELPERS ─────────────────────────────────────────────────

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Helow ama abuur sheet, kuna dar header-ka haddii uusan jirin. */
function getOrCreateSheet_(name, columns) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(columns);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(columns);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Akhri sheet-ka oo dhan, u beddel array of objects ah (header row = keys). */
function readSheetAsObjects_(name, columns) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var rows = values.slice(1);

  return rows.map(function (row) {
    var obj = {};
    header.forEach(function (key, i) {
      obj[key] = row[i] !== undefined ? row[i] : '';
    });
    return obj;
  });
}

/**
 * Ku dar ama cusboonaysii safaf (rows) — haddii idKey (tusaale spfNum) horeyba
 * jiro saf jira, waa la cusboonaysiiyaa; haddii kalena waa lagu daraa saf cusub.
 */
function upsertRecords_(sheetName, columns, records, idKey) {
  var sheet = getOrCreateSheet_(sheetName, columns);
  var lastRow = sheet.getLastRow();
  var idColIndex = columns.indexOf(idKey);
  var existingIds = {};

  // Build a map of idKey -> row number (1-based, including header)
  if (lastRow > 1 && idColIndex !== -1) {
    var idRange = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
    idRange.forEach(function (r, i) {
      if (r[0] !== '' && r[0] !== null) existingIds[String(r[0])] = i + 2; // +2: header + 1-based
    });
  }

  var rowsToAppend = [];
  var written = 0;

  records.forEach(function (rec) {
    rec.updatedAt = new Date().toISOString();
    var rowValues = columns.map(function (col) {
      var v = rec[col];
      if (v === undefined || v === null) return '';
      // Skip large base64 blobs (photos / fingerprint images) to keep the sheet light
      if (typeof v === 'string' && v.length > 3000) return '[binary omitted]';
      return v;
    });

    var idVal = idKey ? String(rec[idKey] || '') : '';
    if (idVal && existingIds[idVal]) {
      // UPDATE existing row
      sheet.getRange(existingIds[idVal], 1, 1, columns.length).setValues([rowValues]);
      written++;
    } else {
      // Queue for append
      rowsToAppend.push(rowValues);
      written++;
    }
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, columns.length)
      .setValues(rowsToAppend);
  }

  return written;
}


// ── OPTIONAL: manual test function (run from the Apps Script editor) ──
function _testDoPost() {
  var fakeEvent = {
    parameter: { type: 'officers' },
    postData: {
      contents: JSON.stringify([
        { spfNum: 'SPF-000999', name: 'Tijaabo Officer', apno: 'AP 999999', status: 'ACTIVE' }
      ])
    }
  };
  var res = doPost(fakeEvent);
  Logger.log(res.getContent());
}
