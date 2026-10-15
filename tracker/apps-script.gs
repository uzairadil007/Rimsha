/**
 * Community Tracker backend (Google Apps Script + Google Sheet).
 *
 * Setup (one time, ~5 minutes):
 *  1. Create a new Google Sheet (any name).
 *  2. Extensions → Apps Script. Delete what's there, paste this whole file, save.
 *  3. In the function dropdown pick "setup" and click Run (allow the permissions).
 *     Then open View → Logs (or Execution log) and copy your ADMIN KEY.
 *  4. Deploy → New deployment → type "Web app".
 *       Execute as: Me    ·    Who has access: Anyone
 *     Click Deploy and copy the Web app URL (ends in /exec).
 *  5. Paste that URL into API_URL at the top of tracker/index.html.
 *
 * Each member's secret token lives only in your private Sheet — never in the
 * public repo — so nobody can update someone else's status without their link.
 */

const SHEET = 'Members';
const MEMBER_COUNT = 20;
const STATUSES = ['Not started', 'In progress', 'Done'];
const HEADERS = ['ID', 'Name', 'Task', 'Status', 'Updated', 'Token'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() < 2) {
    sh.clear();
    sh.appendRow(HEADERS);
    for (let i = 1; i <= MEMBER_COUNT; i++) {
      sh.appendRow([i, 'Member ' + i, '', STATUSES[0], '', newToken_()]);
    }
    sh.setFrozenRows(1);
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) props.setProperty('ADMIN_KEY', newToken_().slice(0, 12));
  Logger.log('ADMIN KEY: ' + props.getProperty('ADMIN_KEY'));
}

function doGet() {
  return json_({ ok: true });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    return json_(handle_(req));
  } catch (err) {
    return json_({ error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function handle_(req) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();

  if (req.action === 'me' || req.action === 'setStatus') {
    const i = findByToken_(rows, req.token);
    if (i < 0) throw new Error('This link is not valid. Ask for a new one.');
    if (req.action === 'setStatus') {
      if (STATUSES.indexOf(req.status) < 0) throw new Error('Unknown status.');
      rows[i][3] = req.status;
      rows[i][4] = new Date();
      sh.getRange(i + 2, 4, 1, 2).setValues([[rows[i][3], rows[i][4]]]);
    }
    return { member: publicRow_(rows[i]) };
  }

  checkAdmin_(req.key);

  if (req.action === 'saveAdmin') {
    (req.members || []).forEach(function (m) {
      const i = rows.findIndex(function (r) { return r[0] === Number(m.id); });
      if (i < 0) return;
      const name = String(m.name || '').slice(0, 80);
      const task = String(m.task || '').slice(0, 2000);
      if (task !== String(rows[i][2])) {
        // A new task starts fresh.
        rows[i][3] = STATUSES[0];
        rows[i][4] = '';
      }
      rows[i][1] = name;
      rows[i][2] = task;
    });
    sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  } else if (req.action === 'resetToken') {
    const i = rows.findIndex(function (r) { return r[0] === Number(req.id); });
    if (i < 0) throw new Error('Unknown member.');
    rows[i][5] = newToken_();
    sh.getRange(i + 2, 6).setValue(rows[i][5]);
  } else if (req.action !== 'admin') {
    throw new Error('Unknown action.');
  }

  return {
    members: rows.map(function (r) {
      const m = publicRow_(r);
      m.token = r[5];
      return m;
    })
  };
}

function checkAdmin_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!real || key !== real) throw new Error('Wrong admin key.');
}

function findByToken_(rows, token) {
  if (typeof token !== 'string' || token.length < 16) return -1;
  return rows.findIndex(function (r) { return r[5] === token; });
}

function publicRow_(r) {
  return {
    id: r[0],
    name: r[1],
    task: r[2],
    status: r[3] || STATUSES[0],
    updated: r[4] instanceof Date ? r[4].toISOString() : ''
  };
}

function newToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
