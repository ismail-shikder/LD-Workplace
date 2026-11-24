/**
 * LD Workplace backend (Google Apps Script)
 * Paste this into: Google Sheet > Extensions > Apps Script.
 * Update steps: paste, Save, run authorize() once, then Deploy > Manage deployments > Edit > New version.
 */

const SHEET_SCHEMA = {
  Batches: ['Batch', 'StartDate', 'EndDate', 'Status', 'PlannedGroup', 'Planned1to1', 'Topics', 'Notes'],
  Agents: ['Batch', 'AgentName', 'EmployeeID', 'SipId', 'Status', 'Feedback', 'AssignedTo'],
  Sessions: ['Date', 'Shift', 'Batch', 'Type', 'Trainer', 'Agent', 'Topic', 'Duration', 'Covered', 'Activities', 'Feedback'],
  Materials: ['Title', 'Topic', 'Type', 'Link', 'Batch', 'Notes'],
  CallReviews: ['CallDate', 'Batch', 'Agent', 'CallType', 'Duration', 'Disposition', 'Remarks', 'Reviewer', 'ReviewDate', 'SipId', 'Number'],
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
const DATA_CACHE_SECONDS = 600;

// The NPS Google Sheet (read only). Month tabs such as "October" and "November" are picked up automatically.
const NPS_SHEET_ID = '1daX7FiNp3Yvlwd_mu8dARJY1uJn3BxDtjPhyJNrY0to';
const NPS_AGENT_TAB = 'Agent details'; // columns: D = SIP, E = employee name
const NPS_ROWS_PER_CALL = 8000;
const MONTH_KEYS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

// Only these people (their full names as written in USERS_TO_CREATE) can add, edit or delete.
// Every other account is view only: it can look at everything and search NPS, but cannot change anything.
const EDITORS = ['Ismail', 'Shawan', 'Nirjhor', 'Shoaib'];
function canEdit_(user) {
  const n = String(user && user.name || '').trim().toLowerCase();
  return EDITORS.some(function (x) { return String(x).trim().toLowerCase() === n; });
}

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

// Run this once after pasting new code. It asks permission to read the NPS sheet.
function authorize() {
  const book = npsBook_();
  const tabs = book.getSheets().map(function (s) { return s.getName(); });
  Logger.log('NPS sheet opened: ' + book.getName());
  Logger.log('Tabs: ' + tabs.join(', '));
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
  users.getRange(1, 1, users.getMaxRows(), 5).setNumberFormat('@');
  users.getRange(1, 1, 1, 5).setValues([['Username', 'Name', 'Salt', 'Hash', 'Status']]).setFontWeight('bold');
  users.setFrozenRows(1);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  bumpVer_();
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
  if (req.action === 'signup') return signup_(req);
  const user = verify_(req.token);
  if (!user) return { ok: false, error: 'auth', message: 'Session expired' };
  switch (req.action) {
    case 'listAll': return listAll_(!!req.fresh);
    case 'list': return { ok: true, rows: readSheet_(checkSheet_(req.sheet)) };
    case 'add':
    case 'update':
    case 'delete':
    case 'addMany':
      if (!canEdit_(user)) return { ok: false, error: 'forbidden', message: 'This account is view only. It cannot add, edit or delete.' };
      return req.action === 'addMany' ? writeMany_(req, user) : write_(req, user);
    case 'npsTabs': return npsTabs_();
    case 'npsRows': return npsRows_(req);
    case 'npsMeta': return npsMeta_();
    case 'npsFetch': return npsFetch_(req);
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
        const status = String(rows[i][4] || '').trim().toLowerCase();
        if (status && status !== 'approved') {
          return { ok: false, error: 'pending', message: status === 'pending' ? 'Your account is waiting for approval by an LD officer.' : 'This account is not active. Ask an LD officer.' };
        }
        cache.remove(failKey);
        const token = Utilities.getUuid() + Utilities.getUuid();
        const user = { u: String(rows[i][0]), name: String(rows[i][1]) };
        cache.put('tok_' + token, JSON.stringify(user), SESSION_SECONDS);
        const out = { ok: true, token: token, name: user.name, canEdit: canEdit_(user), sheetUrl: ss_().getUrl() };
        try { out.data = listAll_(false).data; } catch (err) { /* the site loads the data in a second request instead */ }
        return out;
      }
      break;
    }
  }
  cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 900);
  return { ok: false, error: 'login', message: 'Wrong username or password' };
}

// Anyone can ask for an account. It stays "pending" (cannot sign in) until an LD officer types "approved" in the Status column of the Users sheet.
// New accounts are always view only, because only the names in EDITORS can change data.
const MAX_SIGNUPS_PER_HOUR = 10;
function signup_(req) {
  const username = String(req.username || '').trim().toLowerCase();
  const name = String(req.name || '').replace(/\s+/g, ' ').trim();
  const password = String(req.password || '');
  const bad = function (m) { return { ok: false, error: 'bad_request', message: m }; };
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) return bad('Username must be 3 to 30 letters or numbers.');
  if (name.length < 2 || name.length > 40) return bad('Enter your full name.');
  if (password.length < 8 || password.length > 100) return bad('Password must be at least 8 characters.');
  const reserved = EDITORS.map(function (x) { return String(x).trim().toLowerCase(); })
    .concat(USERS_TO_CREATE.map(function (u) { return String(u[0]).trim().toLowerCase(); }));
  if (reserved.indexOf(username) > -1 || reserved.indexOf(name.toLowerCase()) > -1) return bad('That name or username is reserved. Choose a different one.');
  const cache = CacheService.getScriptCache();
  if (Number(cache.get('signups') || 0) >= MAX_SIGNUPS_PER_HOUR) return { ok: false, error: 'locked', message: 'Too many requests right now. Try again in an hour.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss_().getSheetByName(USERS_SHEET);
    if (!sh) return { ok: false, error: 'server', message: 'Backend is not set up yet. Run setup().' };
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === username) return bad('That username is taken.');
      if (String(rows[i][1]).replace(/\s+/g, ' ').trim().toLowerCase() === name.toLowerCase()) return bad('An account with this name already exists.');
    }
    const salt = Utilities.getUuid();
    sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setNumberFormat('@').setValues([[username, name, salt, hash_(salt, password), 'pending']]);
    cache.put('signups', String(Number(cache.get('signups') || 0) + 1), 3600);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
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

/* ---------- cache (versioned, so any write makes the old copy unused) ---------- */

function ver_() {
  return CacheService.getScriptCache().get('ver') || '0';
}

function bumpVer_() {
  CacheService.getScriptCache().put('ver', String(Date.now()), 21600);
}

function cacheGetBig_(key) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key + '_n') || 0);
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  const parts = cache.getAll(keys);
  let out = '';
  for (let i = 0; i < n; i++) {
    const p = parts[key + '_' + i];
    if (p == null) return null;
    out += p;
  }
  return out;
}

function cachePutBig_(key, str, seconds) {
  const size = 30000; // characters per piece, safely below the 100 KB limit even for Bengali text
  const count = Math.ceil(str.length / size);
  if (count > 60) return;
  const obj = {};
  for (let i = 0; i < count; i++) obj[key + '_' + i] = str.slice(i * size, (i + 1) * size);
  obj[key + '_n'] = String(count);
  try { CacheService.getScriptCache().putAll(obj, seconds); } catch (e) { /* too large: skip caching */ }
}

/* ---------- reading and writing the tracker sheets ---------- */

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

function listAll_(fresh) {
  const key = 'all_' + ver_();
  if (!fresh) {
    const hit = cacheGetBig_(key);
    if (hit) {
      try { return { ok: true, data: JSON.parse(hit) }; } catch (e) { /* rebuild below */ }
    }
  }
  const data = {};
  Object.keys(SHEET_SCHEMA).forEach(function (name) { data[name] = readSheet_(name); });
  cachePutBig_(key, JSON.stringify(data), DATA_CACHE_SECONDS);
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
      bumpVer_();
      return { ok: true, record: rec };
    }

    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
    let rowNum = -1;
    for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(req.id)) { rowNum = i + 2; break; } }
    if (rowNum < 0) return { ok: false, error: 'not_found', message: 'That entry no longer exists. Refresh and try again.' };

    if (req.action === 'delete') {
      const ownerCol = headers.indexOf('CreatedBy');
      const owner = ownerCol > -1 ? String(sh.getRange(rowNum, ownerCol + 1).getValue()).trim().toLowerCase() : '';
      if (!owner || owner !== String(user.name).trim().toLowerCase()) {
        return { ok: false, error: 'forbidden', message: 'Only the person who added this entry can delete it.' };
      }
      sh.deleteRow(rowNum);
      bumpVer_();
      return { ok: true };
    }

    const current = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const ownCol = headers.indexOf('CreatedBy');
    const ownBy = ownCol > -1 ? String(current[ownCol]).trim().toLowerCase() : '';
    if (!ownBy || ownBy !== String(user.name).trim().toLowerCase()) {
      return { ok: false, error: 'forbidden', message: 'Only the person who added this entry can change it.' };
    }
    const rec = {};
    headers.forEach(function (h, c) { rec[h] = cell_(current[c]); });
    SHEET_SCHEMA[name].forEach(function (k) {
      if (req.row && Object.prototype.hasOwnProperty.call(req.row, k)) rec[k] = clean_(req.row[k]);
    });
    rec.UpdatedAt = now;
    const range = sh.getRange(rowNum, 1, 1, headers.length);
    range.setNumberFormat('@');
    range.setValues([rowOf_(headers, rec)]);
    bumpVer_();
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
    bumpVer_();
    return { ok: true, records: recs };
  } finally {
    lock.releaseLock();
  }
}

function clean_(v) {
  return String(v == null ? '' : v).slice(0, MAX_CELL_LENGTH);
}

/* ---------- NPS sheet (read only) ---------- */

function npsBook_() {
  try {
    return SpreadsheetApp.openById(NPS_SHEET_ID);
  } catch (e) {
    throw new Error('Cannot open the NPS sheet. Run authorize() in Apps Script, and make sure the sheet is shared with the account that owns this script. (' + e.message + ')');
  }
}

// "October", "November 2026", "Oct" -> month number 0-11. Anything else (Monthly Report, Agent details) -> -1.
function monthOfTab_(name) {
  const m = /^([a-z]{3})[a-z]*(\s+\d{4})?$/i.exec(String(name).trim());
  if (!m) return -1;
  return MONTH_KEYS.indexOf(m[1].toLowerCase());
}

function npsAgents_(book) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('nps_agents');
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* rebuild below */ }
  }
  const map = {};
  const sh = book.getSheetByName(NPS_AGENT_TAB);
  if (sh && sh.getLastRow() > 1) {
    const vals = sh.getRange(1, 1, sh.getLastRow(), 8).getDisplayValues();
    vals.forEach(function (r) {
      const sip = String(r[3]).trim();
      const name = String(r[4]).replace(/\s+/g, ' ').trim();
      if (/^\d{3,6}$/.test(sip) && name) {
        if (!map[sip]) map[sip] = name;
        else if (map[sip].split(' / ').indexOf(name) < 0) map[sip] += ' / ' + name;
      }
    });
  }
  try { cache.put('nps_agents', JSON.stringify(map), 1800); } catch (e) { /* ignore */ }
  return map;
}

// Used by the website: the list of month tabs, then every NPS call of one tab.
function npsTabs_() {
  const tabs = [];
  npsBook_().getSheets().forEach(function (sh) {
    if (monthOfTab_(sh.getName()) >= 0) tabs.push({ name: sh.getName(), rows: sh.getLastRow() });
  });
  return { ok: true, tabs: tabs };
}

// Each row is [date, time, sip, phone, disposition, ivr, rating, remarks, review, shift, takenBy, action, actionRemarks, actionTaken].
function npsRows_(req) {
  const name = String(req.tab || '');
  const sh = npsBook_().getSheetByName(name);
  const month = monthOfTab_(name);
  if (!sh || month < 0) return { ok: false, error: 'bad_request', message: 'Unknown NPS tab' };
  const last = sh.getLastRow();
  const rows = [];
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, 15).getDisplayValues();
    for (let i = 0; i < vals.length; i++) {
      const rec = npsRow_(vals[i], i + 2, month);
      if (rec) rows.push(rec.slice(1));
    }
  }
  return { ok: true, tab: name, rows: rows };
}

function npsMeta_() {
  const book = npsBook_();
  const tabs = [];
  book.getSheets().forEach(function (sh) {
    if (monthOfTab_(sh.getName()) >= 0) tabs.push({ name: sh.getName(), lastRow: sh.getLastRow() });
  });
  return { ok: true, tabs: tabs, agents: npsAgents_(book) };
}

function npsFetch_(req) {
  const book = npsBook_();
  const name = String(req.tab || '');
  const sh = book.getSheetByName(name);
  const month = monthOfTab_(name);
  if (!sh || month < 0) return { ok: false, error: 'bad_request', message: 'Unknown NPS tab' };
  const lastRow = sh.getLastRow();
  const from = Math.max(2, Math.floor(Number(req.from) || 2));
  if (from > lastRow) return { ok: true, tab: name, lastRow: lastRow, from: from, rows: [], next: 0 };
  const to = Math.min(lastRow, from + NPS_ROWS_PER_CALL - 1);
  const vals = sh.getRange(from, 1, to - from + 1, 15).getDisplayValues();
  const rows = [];
  for (let i = 0; i < vals.length; i++) {
    const rec = npsRow_(vals[i], from + i, month);
    if (rec) rows.push(rec);
  }
  return { ok: true, tab: name, lastRow: lastRow, from: from, to: to, rows: rows, next: to < lastRow ? to + 1 : 0 };
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

function isoDate_(y, mo, d) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return y + '-' + pad2_(mo) + '-' + pad2_(d);
}

// Handles "1-Oct-26", "1 October 2026", "10/1/2026" (month/day, or day/month when the tab month says so) and "2026-10-01".
function parseDate_(s, month) {
  s = String(s || '').trim();
  let m = /^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s,]*(\d{2,4})$/.exec(s);
  if (m) {
    const mi = MONTH_KEYS.indexOf(m[2].slice(0, 3).toLowerCase());
    if (mi < 0) return '';
    const y = Number(m[3]);
    return isoDate_(y < 100 ? 2000 + y : y, mi + 1, Number(m[1]));
  }
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (m) {
    const a = Number(m[1]), b = Number(m[2]);
    let mo = a, d = b;
    if (month >= 0 && b === month + 1 && a !== month + 1) { mo = b; d = a; }
    return isoDate_(Number(m[3]), mo, d);
  }
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return isoDate_(Number(m[1]), Number(m[2]), Number(m[3]));
  return '';
}

function parseTime_(s) {
  const m = /^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?$/.exec(String(s || '').trim());
  if (!m) return '';
  let h = Number(m[1]);
  const ap = m[3] ? m[3].toLowerCase() : '';
  if (ap === 'pm' && h < 12) h += 12;
  if (ap === 'am' && h === 12) h = 0;
  return h > 23 ? '' : pad2_(h) + ':' + m[2];
}

function normPhone_(s) {
  let d = String(s || '').replace(/\D/g, '');
  if (d.length === 13 && d.indexOf('880') === 0) d = '0' + d.slice(3);
  else if (d.length === 10 && d.charAt(0) === '1') d = '0' + d;
  return d;
}

function txt_(v) {
  const s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  return s === '#N/A' ? '' : s;
}

// Sheet columns: A Date, B Time, C SIP ID, D Skill/Campaign (left out on purpose), E Phone, F Disposition,
// G Transfer to IVR, H Rating, I Remarks, J Review, K Shift, L Take over by, M Action, N Remarks, O Action taken.
// Returns [rowNumber, date, time, sip, phone, disposition, ivr, rating, remarks, review, shift, takenBy, action, remarks2, actionTaken]
// or null for headings, day separators and blank rows.
function npsRow_(r, rowNum, month) {
  const date = parseDate_(r[0], month);
  if (!date) return null;
  const phone = normPhone_(r[4]);
  const sip = txt_(r[2]);
  if (!phone && !sip) return null;
  return [rowNum, date, parseTime_(r[1]), sip, phone, txt_(r[5]), txt_(r[6]), txt_(r[7]), txt_(r[8]), txt_(r[9]), txt_(r[10]), txt_(r[11]), txt_(r[12]), txt_(r[13]), txt_(r[14])];
}
