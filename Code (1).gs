/**
 * LD Workplace backend (Google Apps Script)
 * Paste this into: Google Sheet > Extensions > Apps Script.
 * Then run setup() (safe to run again after every update), run createUsers() if needed,
 * and deploy as a Web app (Manage deployments > Edit > New version for updates).
 */

const SHEET_SCHEMA = {
  Batches: ['Batch', 'StartDate', 'EndDate', 'Status', 'PlannedGroup', 'Planned1to1', 'Topics', 'Notes'],
  Agents: ['Batch', 'AgentName', 'EmployeeID', 'Status', 'Feedback'],
  Sessions: ['Date', 'Shift', 'Batch', 'Type', 'Trainer', 'Agent', 'Topic', 'Duration', 'Covered', 'Activities', 'Feedback'],
  Materials: ['Title', 'Topic', 'Type', 'Link', 'Batch', 'Notes'],
  CallReviews: ['CallDate', 'Batch', 'Agent', 'CallType', 'Duration', 'Disposition', 'Remarks', 'Reviewer'],
  Performance: ['Date', 'Batch', 'Agent', 'Category', 'Reason', 'TrainingPlan', 'Status'],
  Handover: ['Batch', 'Agent', 'HandoverDate', 'QAOfficer', 'Method', 'Notes'],
  Settings: ['List', 'Value']
};
const META_COLUMNS = ['CreatedBy', 'CreatedAt', 'UpdatedAt'];
const USERS_SHEET = 'Users';
const SESSION_SECONDS = 21600; // 6 hours (the maximum CacheService allows)
const MAX_FAILED_LOGINS = 5;
const MAX_CELL_LENGTH = 5000;
const MAX_BATCH_ROWS = 200;

// To set one common password for everyone, type it here, run createUsers(), then empty it again.
const SHARED_PASSWORD = '';

// Columns: username, full name, personal password (used only when SHARED_PASSWORD is empty).
const USERS_TO_CREATE = [
  ['ismail', 'Ismail', ''],
  ['shawan', 'Shawan', ''],
  ['nirjhor', 'Nirjhor', ''],
  ['shoaib', 'Shoaib', '']
];

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function wanted_(name) {
  return ['ID'].concat(SHEET_SCHEMA[name], META_COLUMNS);
}

// Creates missing tabs and adds any missing columns at the end. Existing data is never moved or deleted.
function setup() {
  const ss = ss_();
  Object.keys(SHEET_SCHEMA).forEach(function (name) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const lastCol = sh.getLastColumn();
    const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
    wanted_(name).forEach(function (h) { if (headers.indexOf(h) < 0) headers.push(h); });
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const users = ss.getSheetByName(USERS_SHEET) || ss.insertSheet(USERS_SHEET);
  users.getRange(1, 1, users.getMaxRows(), 4).setNumberFormat('@');
  users.getRange(1, 1, 1, 4).setValues([['Username', 'Name', 'Salt', 'Hash']]).setFontWeight('bold');
  users.setFrozenRows(1);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function createUsers() {
  const sh = ss_().getSheetByName(USERS_SHEET);
  if (!sh) throw new Error('Run setup() first.');
  USERS_TO_CREATE.forEach(function (u) {
    if (!(SHARED_PASSWORD || u[2])) throw new Error('Set SHARED_PASSWORD (or a password for each user) at the top of the file first.');
  });
  USERS_TO_CREATE.forEach(function (u) {
    const username = String(u[0]).trim();
    const salt = Utilities.getUuid();
    const row = [username, u[1], salt, hash_(salt, String(SHARED_PASSWORD || u[2]))];
    const values = sh.getDataRange().getValues();
    let target = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim().toLowerCase() === username.toLowerCase()) { target = i + 1; break; }
    }
    if (target < 0) target = sh.getLastRow() + 1;
    sh.getRange(target, 1, 1, 4).setNumberFormat('@').setValues([row]);
  });
}

function hash_(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + password, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function doGet() {
  return json_({ ok: true, app: 'LD Workplace API' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = route_(req);
  } catch (err) {
    out = { ok: false, error: 'server', message: String(err && err.message ? err.message : err) };
  }
  return json_(out);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  if (req.action === 'login') return login_(req);
  const user = verify_(req.token);
  if (!user) return { ok: false, error: 'auth', message: 'Session expired' };
  switch (req.action) {
    case 'listAll': return listAll_();
    case 'list': return { ok: true, rows: readSheet_(checkSheet_(req.sheet)) };
    case 'add':
    case 'update':
    case 'delete': return write_(req, user);
    case 'addMany': return writeMany_(req, user);
    default: return { ok: false, error: 'bad_request', message: 'Unknown action' };
  }
}

function login_(req) {
  const username = String(req.username || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + username.slice(0, 60);
  if (Number(cache.get(failKey) || 0) >= MAX_FAILED_LOGINS) {
    return { ok: false, error: 'locked', message: 'Too many attempts. Try again in 15 minutes.' };
  }
  const sh = ss_().getSheetByName(USERS_SHEET);
  if (!sh) return { ok: false, error: 'server', message: 'Backend is not set up yet. Run setup().' };
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === username) {
      if (hash_(String(rows[i][2]), String(req.password || '')) === String(rows[i][3])) {
        cache.remove(failKey);
        const token = Utilities.getUuid() + Utilities.getUuid();
        const user = { u: String(rows[i][0]), name: String(rows[i][1]) };
        cache.put('tok_' + token, JSON.stringify(user), SESSION_SECONDS);
        return { ok: true, token: token, name: user.name, sheetUrl: ss_().getUrl() };
      }
      break;
    }
  }
  cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 900);
  return { ok: false, error: 'login', message: 'Wrong username or password' };
}

function verify_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('tok_' + String(token));
  return raw ? JSON.parse(raw) : null;
}

function checkSheet_(name) {
  if (!SHEET_SCHEMA[name]) throw new Error('Unknown sheet');
  return name;
}

function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}

function readSheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) {
    if (name === 'Settings') return []; // lets the site work before setup() is re-run
    throw new Error('Sheet "' + name + '" is missing. Run setup().');
  }
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  const out = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    const obj = {};
    headers.forEach(function (h, c) { obj[h] = cell_(values[i][c]); });
    out.push(obj);
  }
  return out;
}

function listAll_() {
  const data = {};
  Object.keys(SHEET_SCHEMA).forEach(function (name) { data[name] = readSheet_(name); });
  return { ok: true, data: data };
}

function sheetFor_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" is missing. Run setup() once, then deploy a new version.');
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  wanted_(name).forEach(function (h) {
    if (headers.indexOf(h) < 0) throw new Error('Sheet "' + name + '" is missing the column "' + h + '". Run setup() once, then deploy a new version.');
  });
  return { sh: sh, headers: headers };
}

function newRecord_(name, row, user, now) {
  const rec = { ID: 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), CreatedBy: user.name, CreatedAt: now, UpdatedAt: '' };
  SHEET_SCHEMA[name].forEach(function (k) { rec[k] = clean_(row && row[k]); });
  return rec;
}

function rowOf_(headers, rec) {
  return headers.map(function (h) { return rec[h] == null ? '' : rec[h]; });
}

function write_(req, user) {
  const name = checkSheet_(req.sheet);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheetFor_(name), sh = s.sh, headers = s.headers;
    const now = new Date().toISOString();

    if (req.action === 'add') {
      const rec = newRecord_(name, req.row, user, now);
      const range = sh.getRange(sh.getLastRow() + 1, 1, 1, headers.length);
      range.setNumberFormat('@');
      range.setValues([rowOf_(headers, rec)]);
      return { ok: true, record: rec };
    }

    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
    let rowNum = -1;
    for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(req.id)) { rowNum = i + 2; break; } }
    if (rowNum < 0) return { ok: false, error: 'not_found', message: 'That entry no longer exists. Refresh and try again.' };

    if (req.action === 'delete') {
      sh.deleteRow(rowNum);
      return { ok: true };
    }

    const current = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const rec = {};
    headers.forEach(function (h, c) { rec[h] = cell_(current[c]); });
    SHEET_SCHEMA[name].forEach(function (k) {
      if (req.row && Object.prototype.hasOwnProperty.call(req.row, k)) rec[k] = clean_(req.row[k]);
    });
    rec.UpdatedAt = now;
    const range = sh.getRange(rowNum, 1, 1, headers.length);
    range.setNumberFormat('@');
    range.setValues([rowOf_(headers, rec)]);
    return { ok: true, record: rec };
  } finally {
    lock.releaseLock();
  }
}

function writeMany_(req, user) {
  const name = checkSheet_(req.sheet);
  const rows = (req.rows || []).slice(0, MAX_BATCH_ROWS);
  if (!rows.length) return { ok: true, records: [] };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheetFor_(name), sh = s.sh, headers = s.headers;
    const now = new Date().toISOString();
    const recs = rows.map(function (r) { return newRecord_(name, r, user, now); });
    const range = sh.getRange(sh.getLastRow() + 1, 1, recs.length, headers.length);
    range.setNumberFormat('@');
    range.setValues(recs.map(function (rec) { return rowOf_(headers, rec); }));
    return { ok: true, records: recs };
  } finally {
    lock.releaseLock();
  }
}

function clean_(v) {
  return String(v == null ? '' : v).slice(0, MAX_CELL_LENGTH);
}
