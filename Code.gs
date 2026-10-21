// Grade 5 Math SOL Practice: serves the site, signs in students (username + PIN) and teachers
// (username + password), saves every answer, and tracks mastery by standard.

// The 14 Grade 5 standards in the 2023 Virginia Mathematics SOL.
const STANDARD_CODES = ['5.NS.1', '5.NS.2', '5.CE.1', '5.CE.2', '5.CE.3', '5.CE.4', '5.MG.1', '5.MG.2', '5.MG.3',
  '5.PS.1', '5.PS.2', '5.PS.3', '5.PFA.1', '5.PFA.2'];
const RESULT_HEADERS = ['Timestamp', 'Username', 'Student name', 'Mode', 'Standard', 'Correct on 1st try', 'Correct by 2nd try',
  'Tries', 'Hints used', 'Question type', 'Question', 'Student answer', 'Correct answer', 'Sub-standard'];

// Mastery rule: each student's last 10 questions on a standard, first try only.
const LAST_N = 10;
const MASTERED_AT = 0.8;     // 80% or more, with at least MIN_DONE questions
const ALMOST_AT = 0.6;       // 60% to 79%
const MIN_DONE = 5;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Grade 5 Math SOL Practice')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Run this once from the editor: choose "setup" in the toolbar and click Run.
// It creates the spreadsheet and a first teacher sign-in, and prints both in the Execution log.
function setup() {
  const ss = getSpreadsheet_();
  Logger.log('Spreadsheet: ' + ss.getUrl());
  const t = ss.getSheetByName('Teachers').getDataRange().getDisplayValues()[1];
  if (t) Logger.log('Teacher sign-in: username "' + t[0] + '", password "' + t[1] + '". Change them in the Teachers tab.');
}

/* ---------- Student calls ---------- */

function login(username, pin) {
  const s = checkStudent_(username, pin);
  return { username: s.username, name: s.name, mastery: masteryMap_(readResults_())[s.username.toLowerCase()] || {} };
}

function logResults(username, pin, rows) {
  const s = checkStudent_(username, pin);
  if (!Array.isArray(rows) || !rows.length || rows.length > 50) throw new Error('Nothing to save.');
  const now = new Date();
  const clean = rows.map(r => [now, s.username, s.name].concat(RESULT_HEADERS.slice(3).map((_, i) => safe_(r[i]))));
  return withLock_(() => {
    const sh = getSpreadsheet_().getSheetByName('Results');
    sh.getRange(sh.getLastRow() + 1, 1, clean.length, RESULT_HEADERS.length).setValues(clean);
    const m = masteryMap_(readResults_())[s.username.toLowerCase()] || {};
    writeMasteryRow_(s, m);
    return m;
  });
}

/* ---------- Teacher calls ---------- */

function teacherLogin(username, password) {
  const t = checkTeacher_(username, password);
  return { username: t.username, name: t.name, cls: classData_(t) };
}

function getClass(username, password) {
  return classData_(checkTeacher_(username, password));
}

// Adds a student, or updates the name and PIN of one already in this teacher's class.
function saveStudent(username, password, student, isNew) {
  const t = checkTeacher_(username, password);
  const u = String(student && student.username || '').trim(), pin = String(student && student.pin || '').trim(), name = String(student && student.name || '').trim();
  if (!/^[A-Za-z0-9._-]{2,30}$/.test(u)) throw new Error('Usernames need 2 to 30 letters, numbers, dots, dashes, or underscores, with no spaces.');
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PINs need 4 to 8 digits.');
  if (!name || name.length > 60) throw new Error('Type the student’s name (60 characters or fewer).');
  return withLock_(() => {
    const sh = getSpreadsheet_().getSheetByName('Students'), rows = sh.getDataRange().getDisplayValues();
    const at = rows.findIndex((r, i) => i > 0 && String(r[0]).trim().toLowerCase() === u.toLowerCase());
    if (isNew) {
      if (at > 0) throw new Error('The username "' + u + '" is already taken. Choose another one.');
      sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[u, pin, safe_(name), t.username]]);
    } else {
      if (at < 0 || !visibleTo_(rows[at], t)) throw new Error('That student is not in your class.');
      sh.getRange(at + 1, 2, 1, 2).setValues([[pin, safe_(name)]]);
    }
    return classData_(t);
  });
}

function deleteStudent(username, password, studentUsername) {
  const t = checkTeacher_(username, password), u = String(studentUsername || '').trim().toLowerCase();
  return withLock_(() => {
    const ss = getSpreadsheet_(), sh = ss.getSheetByName('Students'), rows = sh.getDataRange().getDisplayValues();
    const at = rows.findIndex((r, i) => i > 0 && String(r[0]).trim().toLowerCase() === u);
    if (at < 0 || !visibleTo_(rows[at], t)) throw new Error('That student is not in your class.');
    sh.deleteRow(at + 1);  // their answers stay in the Results tab
    const m = ss.getSheetByName('Mastery'), mr = m.getDataRange().getDisplayValues().findIndex((r, i) => i > 0 && String(r[1]).toLowerCase() === u);
    if (mr > 0) m.deleteRow(mr + 1);
    return classData_(t);
  });
}

// The 25 most recent answers from one student in this teacher's class.
function studentResults(username, password, studentUsername) {
  const t = checkTeacher_(username, password), u = String(studentUsername || '').trim().toLowerCase();
  const row = studentRows_().find(r => String(r[0]).trim().toLowerCase() === u);
  if (!row || !visibleTo_(row, t)) throw new Error('That student is not in your class.');
  return readResults_().filter(r => String(r[1]).toLowerCase() === u).slice(-25).reverse()
    .map(r => ({ when: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), mode: r[3], std: r[4], ok: Number(r[5]) === 1, okAny: Number(r[6]) === 1, q: r[10], a: r[11], c: r[12] }));
}

/* ---------- Checking sign-ins ---------- */

function checkStudent_(username, pin) {
  const u = String(username || '').trim().toLowerCase(), p = String(pin || '').trim();
  if (!u || !p) throw new Error('Type your username and PIN.');
  limit_('s_' + u);
  const row = studentRows_().find(r => String(r[0]).trim().toLowerCase() === u && String(r[1]).trim() === p);
  if (!row) { fail_('s_' + u); throw new Error("That username and PIN don't match. Check them and try again."); }
  CacheService.getScriptCache().remove('s_' + u);
  return { username: String(row[0]).trim(), name: String(row[2] || row[0]).trim() };
}

function checkTeacher_(username, password) {
  const u = String(username || '').trim().toLowerCase(), p = String(password || '');
  if (!u || !p) throw new Error('Type your username and password.');
  limit_('t_' + u);
  const row = getSpreadsheet_().getSheetByName('Teachers').getDataRange().getDisplayValues().slice(1)
    .find(r => String(r[0]).trim().toLowerCase() === u && String(r[1]) === p && p.length > 0);
  if (!row) { fail_('t_' + u); throw new Error("That username and password don't match. Check them and try again."); }
  CacheService.getScriptCache().remove('t_' + u);
  return { username: String(row[0]).trim(), name: String(row[2] || row[0]).trim() };
}

function limit_(key) {
  if (Number(CacheService.getScriptCache().get(key) || 0) >= 8) throw new Error('Too many tries. Wait 10 minutes and try again.');
}
function fail_(key) {
  const c = CacheService.getScriptCache();
  c.put(key, String(Number(c.get(key) || 0) + 1), 600);
}

// A student belongs to the teacher in the Teacher column; a blank Teacher column means every teacher can see them.
function visibleTo_(row, t) {
  const owner = String(row[3] || '').trim().toLowerCase();
  return !owner || owner === t.username.toLowerCase();
}

/* ---------- Mastery ---------- */

function studentRows_() {
  return getSpreadsheet_().getSheetByName('Students').getDataRange().getDisplayValues().slice(1).filter(r => String(r[0]).trim());
}

function readResults_() {
  const sh = getSpreadsheet_().getSheetByName('Results'), last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, RESULT_HEADERS.length).getValues();
}

// { username: { standard: { done, right, level } } } from the last LAST_N first-try answers.
function masteryMap_(results) {
  const raw = {}, out = {};
  results.forEach(r => {
    const u = String(r[1]).toLowerCase(), code = r[4];
    ((raw[u] = raw[u] || {})[code] = raw[u][code] || []).push(Number(r[5]) === 1 ? 1 : 0);
  });
  Object.keys(raw).forEach(u => {
    out[u] = {};
    Object.keys(raw[u]).forEach(code => {
      const recent = raw[u][code].slice(-LAST_N), done = recent.length, right = recent.reduce((a, b) => a + b, 0), pct = right / done;
      out[u][code] = { done: done, right: right, level: pct >= MASTERED_AT && done >= MIN_DONE ? 'Mastered' : pct >= ALMOST_AT ? 'Almost there' : 'Practicing' };
    });
  });
  return out;
}

function classData_(t) {
  const results = readResults_(), map = masteryMap_(results), lastSeen = {};
  results.forEach(r => { lastSeen[String(r[1]).toLowerCase()] = r[0] instanceof Date ? r[0].toISOString() : String(r[0]); });
  const students = studentRows_().filter(r => visibleTo_(r, t)).map(r => {
    const u = String(r[0]).trim();
    return { username: u, pin: String(r[1]).trim(), name: String(r[2] || u).trim(), mastery: map[u.toLowerCase()] || {}, last: lastSeen[u.toLowerCase()] || '' };
  }).sort((a, b) => a.name.localeCompare(b.name));
  return { students: students, codes: STANDARD_CODES };
}

function writeMasteryRow_(student, m) {
  const sh = getSpreadsheet_().getSheetByName('Mastery');
  const row = [student.name, student.username, new Date()].concat(STANDARD_CODES.map(c => m[c] ? m[c].level + ' · ' + m[c].right + '/' + m[c].done : 'Not started'));
  const names = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).toLowerCase()) : [];
  const at = names.indexOf(student.username.toLowerCase());
  sh.getRange(at >= 0 ? at + 2 : sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

/* ---------- Spreadsheet setup ---------- */

function getSpreadsheet_() {
  // If this script was opened from a Sheet (Extensions > Apps Script), use that Sheet.
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return prepare_(active);
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('RESULTS_ID');
  if (id) {
    try { return prepare_(SpreadsheetApp.openById(id)); } catch (e) { /* deleted: make a new one */ }
  }
  const ss = SpreadsheetApp.create('Grade 5 Math SOL Practice - Students and Mastery');
  props.setProperty('RESULTS_ID', ss.getId());
  return prepare_(ss);
}

function prepare_(ss) {
  if (!ss.getSheetByName('Teachers')) {
    const blank = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0] : null;
    const t = blank ? blank.setName('Teachers') : ss.insertSheet('Teachers', 0);
    t.getRange(1, 1, 1, 3).setValues([['Username', 'Password', 'Teacher name']]).setFontWeight('bold');
    t.getRange(2, 1, 1, 3).setValues([['teacher', randomPassword_(), 'Teacher']]);
    t.getRange('E1').setValue('One row per teacher. Change the username, password, and name as you like.');
    t.setFrozenRows(1);
  }
  if (!ss.getSheetByName('Students')) {
    const st = ss.insertSheet('Students', 1);
    st.getRange('B:B').setNumberFormat('@');  // keep PINs like 0420 exactly as typed
    st.getRange(1, 1, 1, 4).setValues([['Username', 'PIN', 'Student name', 'Teacher']]).setFontWeight('bold');
    st.getRange('F1').setValue('Teachers can add students from the website. Teacher = the teacher’s username (blank = every teacher sees the student).');
    st.setFrozenRows(1);
  }
  if (!ss.getSheetByName('Results')) {
    const r = ss.insertSheet('Results');
    r.setFrozenRows(1);
  }
  ss.getSheetByName('Results').getRange(1, 1, 1, RESULT_HEADERS.length).setValues([RESULT_HEADERS]).setFontWeight('bold');
  if (!ss.getSheetByName('Mastery')) {
    const m = ss.insertSheet('Mastery');
    m.getRange(1, 1, 1, 3 + STANDARD_CODES.length).setValues([['Student', 'Username', 'Last active'].concat(STANDARD_CODES)]).setFontWeight('bold');
    m.setFrozenRows(1); m.setFrozenColumns(1);
    const grid = m.getRange(2, 4, 400, STANDARD_CODES.length), rule = (text, color) => SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(text).setBackground(color).setRanges([grid]).build();
    m.setConditionalFormatRules([rule('Mastered', '#c8ecd4'), rule('Almost', '#fdf0b8'), rule('Practicing', '#f8d3c5'), rule('Not started', '#eeeeee')]);
  }
  // keep the Mastery headers in step with the current list of standards
  const mh = ss.getSheetByName('Mastery'), heads = ['Student', 'Username', 'Last active'].concat(STANDARD_CODES);
  mh.getRange(1, 1, 1, heads.length).setValues([heads]).setFontWeight('bold');
  mh.getRange(1, heads.length + 1, 1, 3).setValues([['', '', '']]);
  return ss;
}

function randomPassword_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 10; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// Keep cell values plain text so nothing typed can run as a formula.
function safe_(v) {
  if (typeof v === 'number') return v;
  let s = String(v == null ? '' : v).slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
