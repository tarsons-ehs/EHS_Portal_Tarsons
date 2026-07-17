/**
 * TARSONS HSE Portal — Google Sheets backend
 * ---------------------------------------------------------
 * Deploy this as a Web App (see instructions in chat / README).
 * It gives the HTML portal a shared, network-visible database:
 *   - Every LocalDB "key" (users, ptw, hira, ppe, gallery, ...) gets
 *     its own sheet tab, created automatically the first time it's saved.
 *   - Cell A1/B1 hold a JSON blob of the exact data (this is what the
 *     app reads back — guaranteed to round-trip perfectly).
 *   - Starting at row 4, the same data is also written out as a normal
 *     table (one column per field) purely so a human opening the sheet
 *     can read it.
 */

const JSON_MARKER = '__json__';

function doGet(e) {
  try {
    const action = (e.parameter.action || '').toLowerCase();
    if (action === 'getall') {
      return jsonOut({ status: 'ok', data: getAllData() });
    }
    if (action === 'get') {
      return jsonOut({ status: 'ok', data: getKey(e.parameter.key) });
    }
    return jsonOut({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'update') {
      saveKey(body.key, body.value);
      return jsonOut({ status: 'ok' });
    }
    if (body.action === 'delete') {
      deleteKeySheet(body.key);
      return jsonOut({ status: 'ok' });
    }
    return jsonOut({ status: 'error', message: 'Unknown action: ' + body.action });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Sheet names can't contain []*/\?: and are capped at 100 chars.
function sanitizeSheetName(name) {
  return String(name).replace(/[\[\]\*\/\\\?:]/g, '_').substring(0, 90) || 'data';
}

function saveKey(key, value) {
  const ss = getSS();
  const sheetName = sanitizeSheetName(key);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();

  // Row 1-2: the authoritative JSON blob the app actually reads back.
  const json = JSON.stringify(value);
  sheet.getRange(1, 1, 2, 2).setValues([
    [JSON_MARKER, json],
    ['updatedAt', new Date().toISOString()]
  ]);

  // Row 4+: human-readable table, best-effort, for arrays of objects only.
  if (Array.isArray(value) && value.length && value.every(v => v && typeof v === 'object' && !Array.isArray(v))) {
    const headers = [];
    value.forEach(row => Object.keys(row).forEach(k => {
      if (headers.indexOf(k) === -1) headers.push(k);
    }));
    if (headers.length) {
      const rows = value.map(row => headers.map(h => {
        const v = row[h];
        if (v === undefined || v === null) return '';
        if (typeof v === 'object') return JSON.stringify(v).substring(0, 49000);
        return String(v).substring(0, 49000);
      }));
      sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(5, 1, rows.length, headers.length).setValues(rows);
    }
  }

  try { sheet.autoResizeColumns(1, Math.max(2, sheet.getLastColumn())); } catch (e) {}
}

function deleteKeySheet(key) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sanitizeSheetName(key));
  if (sheet) ss.deleteSheet(sheet);
}

function getKey(key) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sanitizeSheetName(key));
  if (!sheet) return null;
  const raw = sheet.getRange(1, 2).getValue();
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function getAllData() {
  const ss = getSS();
  const data = {};
  ss.getSheets().forEach(sheet => {
    const marker = sheet.getRange(1, 1).getValue();
    if (marker === JSON_MARKER) {
      const raw = sheet.getRange(1, 2).getValue();
      try { data[sheet.getName()] = JSON.parse(raw); } catch (e) { /* skip corrupt tab */ }
    }
  });
  return data;
}

/**
 * --- Daily overdue-alert email ---
 * 1. Set ALERT_EMAIL below to where reminders should go (comma-separate
 *    multiple addresses, e.g. "a@x.com,b@x.com").
 * 2. In the Apps Script editor: Triggers (clock icon, left sidebar) →
 *    Add Trigger → function: sendOverdueAlertEmail → Event source:
 *    Time-driven → Day timer → pick an hour → Save.
 * That's it — it'll run once a day and only emails if something is
 * actually due soon or overdue, so you won't get spammed on quiet days.
 */
const ALERT_EMAIL = "PASTE_YOUR_EMAIL_HERE";

function sendOverdueAlertEmail() {
  if (!ALERT_EMAIL || ALERT_EMAIL.indexOf('PASTE_YOUR') === 0) return; // not configured yet
  const today = new Date();
  const daysUntil = (d) => (new Date(d) - today) / (1000 * 60 * 60 * 24);

  const extinguishers = getKey('extinguishers') || [];
  const legal = getKey('legal') || [];
  const capa = getKey('capa') || [];

  const extDue = extinguishers.filter(x => x.due && daysUntil(x.due) <= 30);
  const legalDue = legal.filter(x => x.expiry && daysUntil(x.expiry) <= 30);
  const capaOverdue = capa.filter(x => x.status !== 'Closed' && x.due && daysUntil(x.due) < 0);

  if (!extDue.length && !legalDue.length && !capaOverdue.length) return; // nothing to report today

  let body = 'TARSONS HSE Portal — Daily Alert Summary\n\n';
  if (extDue.length) {
    body += 'FIRE EXTINGUISHERS DUE / OVERDUE:\n';
    extDue.forEach(x => body += `  - ${x.extID} (${x.factory}) — due ${x.due}\n`);
    body += '\n';
  }
  if (legalDue.length) {
    body += 'LEGAL / STATUTORY LICENSES EXPIRING:\n';
    legalDue.forEach(x => body += `  - ${x.license} (${x.factory}) — expires ${x.expiry}\n`);
    body += '\n';
  }
  if (capaOverdue.length) {
    body += 'OVERDUE CAPA ACTIONS:\n';
    capaOverdue.forEach(x => body += `  - ${x.desc} — owner: ${x.owner} (${x.factory}), was due ${x.due}\n`);
    body += '\n';
  }
  body += 'Open the portal to review and update these.';

  MailApp.sendEmail(ALERT_EMAIL, 'HSE Portal: items due or overdue', body);
}
