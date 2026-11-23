/**
 * Cloud Gallery - Google Apps Script backend (rebuilt)
 *
 * Deploy as: Web app -> Execute as "Me" -> Who has access "Anyone".
 * Run setup() once from the editor to authorize and create the database.
 *
 * Storage layout
 *  - Google Sheet "Photo Gallery DB"
 *      Users  : one row per user (profile, password hash, prefs, folders)
 *      Shares : public share links
 *      <uid>  : one sheet per user, one row per photo/video
 *  - Google Drive folder "Photo Gallery Files" / "user_<uid>" : the actual files
 *
 * Videos are uploaded in chunks (resumable upload relayed through this script),
 * because a single Apps Script request cannot carry a 100 MB file.
 */

/* ============================ CONFIG ============================ */
const ADMIN_EMAIL = 'rizzuid1@gmail.com';
const APP_NAME = 'Cloud Gallery';
const DEFAULT_LIMIT_GB = 5;                    // default storage per user
const MAX_ITEMS_PER_USER = 1000;               // hidden from users
const MAX_VIDEO_BYTES = 100 * 1024 * 1024;     // 100 MB per video
const MAX_IMAGE_BYTES = 25 * 1024 * 1024;      // single-request image upload
const TRASH_DAYS = 30;
const SESSION_DAYS = 30;                       // stay logged in on a device
const HASH_ROUNDS = 500;
const MAX_FAILS = 5;                           // wrong passwords allowed ...
const FAIL_WINDOW_SEC = 900;                   // ... per 15 minutes per e-mail
const RESET_MINUTES = 60;                      // how long an e-mailed temporary password stays valid
const RESET_MAX_MAILS = 3;                     // forgot-password e-mails per address per hour
/* ================================================================ */

const RESERVED = ['my drive', 'camera', 'favorites', 'trash'];
const DEFAULT_FOLDERS = ['My Drive', 'Camera', 'Favorites', 'Trash'];
const UHEAD = ['email', 'name', 'uid', 'salt', 'hash', 'verified', 'vcode', 'vexp', 'rcode', 'rexp', 'limitGB', 'used', 'blocked', 'prefs', 'folders', 'created', 'driveId'];
const U = {}; UHEAD.forEach((k, i) => U[k] = i);
const RHEAD = ['id', 'folder', 'ts', 'size', 'name', 'kind', 'dur', 'fav', 'trashed'];
const R = {}; RHEAD.forEach((k, i) => R[k] = i);
const SHEAD = ['token', 'uid', 'type', 'target', 'created', 'active'];
const SESHEAD = ['tokenHash', 'uid', 'expires'];
const UPLOAD_HOST = 'https://www.googleapis.com/upload/drive/v3/files?';

class ApiError extends Error { constructor(msg, code) { super(msg); this.code = code || ''; } }
const fail_ = (m, c) => { throw new ApiError(m, c); };
const N_ = v => { const n = Number(v); return isFinite(n) ? n : 0; };

/* ============================ SETUP ============================ */
function setup() {
  const ss = db_(); rootFolder_(); sessionsSheet_();
  Logger.log('Database ready: ' + ss.getUrl());
  Logger.log('Next: run authorizeMail once (click Allow), then Deploy > New deployment > Web app > Execute as Me > Anyone, and copy the /exec URL into index.html (API_URL).');
}
/* Run once from the editor: asks for the e-mail permission (needed for Forgot password) and sends a test mail. */
function authorizeMail() {
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: APP_NAME + ' - test e-mail', body: 'If you can read this, e-mail sending works.' });
  Logger.log('Test e-mail sent to ' + ADMIN_EMAIL + '. Remaining daily quota: ' + MailApp.getRemainingDailyQuota());
}

let _ss = null;
function db_() {
  if (_ss) return _ss;
  const p = PropertiesService.getScriptProperties(); const id = p.getProperty('SHEET_ID');
  if (id) { try { _ss = SpreadsheetApp.openById(id); return _ss; } catch (e) { /* recreate below */ } }
  const ss = SpreadsheetApp.create('Photo Gallery DB');
  const u = ss.getSheets()[0]; u.setName('Users');
  u.getRange(1, 1, 1, UHEAD.length).setValues([UHEAD]); u.getRange(1, 1, 1000, UHEAD.length).setNumberFormat('@');
  const s = ss.insertSheet('Shares');
  s.getRange(1, 1, 1, SHEAD.length).setValues([SHEAD]); s.getRange(1, 1, 5000, SHEAD.length).setNumberFormat('@');
  p.setProperty('SHEET_ID', ss.getId()); _ss = ss; return ss;
}
function rootFolder_() {
  const p = PropertiesService.getScriptProperties(); const id = p.getProperty('ROOT_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreate */ } }
  const f = DriveApp.createFolder('Photo Gallery Files'); p.setProperty('ROOT_ID', f.getId()); return f;
}
/* ============================ HTTP ============================ */
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  try {
    const t = e && e.parameter && e.parameter.share;
    if (!t) return out_({ ok: true, msg: APP_NAME + ' API is running' });
    return out_(locked_(() => sharedView_(String(t))));
  } catch (err) { return out_({ ok: false, error: err.message || String(err) }); }
}

const ROUTES = {
  signup: signup_, login: login_, forgotPassword: forgotPassword_, logout: logout_,
  list: list_, updateName: updateName_, changePassword: changePassword_, deleteAccount: deleteAccount_,
  createFolder: createFolder_, renameFolder: renameFolder_, deleteFolder: deleteFolder_, setCover: setCover_, setOrder: setOrder_,
  upload: upload_, vStart: vStart_, vChunk: vChunk_,
  remove: remove_, move: move_, favorite: favorite_, emptyTrash: emptyTrash_, info: info_, recount: recount_,
  share: share_, unshare: unshare_,
  adminUsers: adminUsers_, adminSet: adminSet_, adminRecount: adminRecount_, adminPhotos: adminPhotos_, adminDelete: adminDelete_
};
const PUBLIC = { signup: 1, login: 1, forgotPassword: 1, logout: 1 };
const UNLOCKED = { vStart: 1, vChunk: 1 };   // slow network relays; they lock internally only where needed

function doPost(e) {
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    const a = String(d.action || ''); const fn = ROUTES[a];
    if (!fn) fail_('Unknown action');
    const run = () => PUBLIC[a] ? fn(d) : fn(d, auth_(d.token));
    return out_(UNLOCKED[a] ? run() : locked_(run));
  } catch (err) {
    return out_({ ok: false, error: err.message || String(err), code: err.code || '' });
  }
}
function locked_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(30000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ============================ Users & auth ============================ */
const usersSheet_ = () => db_().getSheetByName('Users');
function findUser_(by, value) {
  const sh = usersSheet_(), last = sh.getLastRow(); if (last < 2) return null;
  const vals = sh.getRange(2, U[by] + 1, last - 1, 1).getValues(), v = String(value).toLowerCase();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).toLowerCase() === v) {
    const row = i + 2; return { row, v: sh.getRange(row, 1, 1, UHEAD.length).getValues()[0] };
  }
  return null;
}
const saveUser_ = u => usersSheet_().getRange(u.row, 1, 1, UHEAD.length).setValues([u.v]);
const uid_ = u => String(u.v[U.uid]);
const isAdmin_ = u => String(u.v[U.email]).toLowerCase() === ADMIN_EMAIL.toLowerCase();
const prefsOf_ = u => { try { return JSON.parse(u.v[U.prefs] || '{}') || {}; } catch (e) { return {}; } };
const setPrefs_ = (u, p) => { u.v[U.prefs] = JSON.stringify(p); saveUser_(u); };
const customFolders_ = u => { try { const a = JSON.parse(u.v[U.folders] || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } };
const limitBytes_ = u => (N_(u.v[U.limitGB]) || DEFAULT_LIMIT_GB) * 1073741824;

const hex_ = bytes => bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
function hashPw_(pw, salt) {
  let h = salt + pw;
  for (let i = 0; i < HASH_ROUNDS; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return h;
}
const sha256_ = str => hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8));
const sameStr_ = (a, b) => { a = String(a); b = String(b); if (a.length !== b.length) return false; let d = 0; for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i); return d === 0; };
const newId_ = () => Utilities.getUuid().replace(/-/g, '');
const pwOk_ = (u, pw) => sameStr_(hashPw_(String(pw || ''), u.v[U.salt]), u.v[U.hash]);

/* ---- sessions: only a hash of the token is stored; logging out deletes it ---- */
function sessionsSheet_() {
  let sh = db_().getSheetByName('Sessions');
  if (!sh) { sh = db_().insertSheet('Sessions'); sh.getRange(1, 1, 1, SESHEAD.length).setValues([SESHEAD]); sh.getRange(1, 1, 5000, SESHEAD.length).setNumberFormat('@'); }
  return sh;
}
function newSession_(u) {
  const sh = sessionsSheet_(), last = sh.getLastRow(), now = Date.now();
  if (last > 1) {                                                   // clean expired sessions
    const vals = sh.getRange(2, 1, last - 1, SESHEAD.length).getValues(), keep = vals.filter(r => N_(r[2]) > now);
    if (keep.length !== vals.length) { sh.getRange(2, 1, last - 1, SESHEAD.length).clearContent(); if (keep.length) sh.getRange(2, 1, keep.length, SESHEAD.length).setValues(keep); }
  }
  const token = newId_() + newId_();
  sh.appendRow([sha256_(token), uid_(u), now + SESSION_DAYS * 864e5]);
  return token;
}
function dropSessions_(match) {                                     // match(row) -> true = delete
  const sh = sessionsSheet_(), last = sh.getLastRow(); if (last < 2) return;
  const vals = sh.getRange(2, 1, last - 1, SESHEAD.length).getValues(), keep = vals.filter(r => !match(r));
  if (keep.length === vals.length) return;
  sh.getRange(2, 1, last - 1, SESHEAD.length).clearContent(); if (keep.length) sh.getRange(2, 1, keep.length, SESHEAD.length).setValues(keep);
}
function auth_(token) {
  const bad = () => fail_('Your session has expired. Please log in again.', 'AUTH');
  if (!token) bad();
  const th = sha256_(String(token)), cache = CacheService.getScriptCache();
  let uid = cache.get('sess_' + th);
  if (!uid) {
    const sh = sessionsSheet_(), last = sh.getLastRow(); if (last < 2) bad();
    const row = sh.getRange(2, 1, last - 1, SESHEAD.length).getValues().find(r => String(r[0]) === th);
    if (!row || N_(row[2]) < Date.now()) bad();
    uid = String(row[1]); cache.put('sess_' + th, uid, 300);
  }
  const u = findUser_('uid', uid); if (!u) bad();
  if (N_(u.v[U.blocked])) fail_('This account has been blocked. Please contact support.', 'AUTH');
  return u;
}
function logout_(d) {
  if (!d.token) return { ok: true };
  const th = sha256_(String(d.token)); CacheService.getScriptCache().remove('sess_' + th);
  dropSessions_(r => String(r[0]) === th); return { ok: true };
}

/* ---- login lock-out: 5 wrong passwords = 15 minutes break ---- */
const failKey_ = email => 'fail_' + sha256_(email).slice(0, 40);
function checkThrottle_(email) { if (N_(CacheService.getScriptCache().get(failKey_(email))) >= MAX_FAILS) fail_('Too many wrong attempts. Please try again after 15 minutes.'); }
function addFail_(email) { const c = CacheService.getScriptCache(), k = failKey_(email); c.put(k, String(N_(c.get(k)) + 1), FAIL_WINDOW_SEC); }
function clearFail_(email) { CacheService.getScriptCache().remove(failKey_(email)); }

function publicUser_(u, rows) {
  rows = rows || readRows_(uid_(u));
  return { name: u.v[U.name], email: u.v[U.email], verified: true, admin: isAdmin_(u), used: usedOf_(rows), limit: limitBytes_(u) };
}
const usedOf_ = rows => rows.reduce((a, r) => a + N_(r[R.size]), 0);

function state_(u) {
  const rows = readRows_(uid_(u));
  const names = DEFAULT_FOLDERS.slice(); customFolders_(u).forEach(n => { if (!names.some(x => x.toLowerCase() === n.toLowerCase())) names.push(n); });
  const map = {}; names.forEach(n => map[n.toLowerCase()] = { name: n, photos: [] });
  const meta = {};
  rows.forEach(r => {
    const id = r[R.id], key = String(r[R.folder]).toLowerCase();
    (map[key] || map['my drive']).photos.push(id);
    if (N_(r[R.fav]) && key !== 'trash') map['favorites'].photos.push(id);
    meta[id] = [N_(r[R.ts]), N_(r[R.size]), r[R.name], r[R.kind] === 'v' ? 'v' : 'i', N_(r[R.dur])];
  });
  return { user: publicUser_(u, rows), folders: names.map(n => map[n.toLowerCase()]), prefs: prefsOf_(u), meta };
}

function signup_(d) {
  const email = String(d.email || '').trim().toLowerCase(), name = String(d.name || '').trim().replace(/\s+/g, ' ').slice(0, 60), pw = String(d.password || '');
  if (!name) fail_('Please enter your name.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail_('Please enter a valid email.');
  if (pw.length < 6) fail_('Password must be at least 6 characters.');
  if (pw.length > 100) fail_('Password is too long.');
  if (findUser_('email', email)) fail_('This email is already registered. Please log in.');
  const uid = 'u' + newId_().slice(0, 12), salt = newId_();
  const folder = rootFolder_().createFolder('user_' + uid);
  const v = new Array(UHEAD.length).fill('');
  v[U.email] = email; v[U.name] = name; v[U.uid] = uid; v[U.salt] = salt; v[U.hash] = hashPw_(pw, salt);
  v[U.verified] = 1; v[U.limitGB] = DEFAULT_LIMIT_GB; v[U.used] = 0; v[U.blocked] = 0; v[U.prefs] = '{}'; v[U.folders] = '[]';
  v[U.created] = Date.now(); v[U.driveId] = folder.getId();
  const sh = usersSheet_(); sh.appendRow(v);
  const us = db_().insertSheet(uid); us.getRange(1, 1, 1, RHEAD.length).setValues([RHEAD]); us.getRange(1, 1, 3000, RHEAD.length).setNumberFormat('@');
  const u = { row: sh.getLastRow(), v };
  return Object.assign({ ok: true, token: newSession_(u) }, state_(u));
}
function login_(d) {
  const email = String(d.email || '').trim().toLowerCase(), pw = String(d.password || '');
  if (!email || !pw) fail_('Please enter email and password.');
  checkThrottle_(email);
  const u = findUser_('email', email), cache = CacheService.getScriptCache();
  let ok = !!u && pwOk_(u, pw);
  if (!ok && u) {                                                   // maybe the temporary password that was e-mailed
    let t = null; try { t = JSON.parse(cache.get('reset_' + sha256_(email).slice(0, 40)) || 'null'); } catch (e) {}
    if (t && sameStr_(hashPw_(pw, t.salt), t.hash)) {
      ok = true;                                                    // it becomes the real password now
      const salt = newId_(); u.v[U.salt] = salt; u.v[U.hash] = hashPw_(pw, salt); saveUser_(u);
      cache.remove('reset_' + sha256_(email).slice(0, 40));
    }
  }
  if (!ok) { addFail_(email); fail_('Wrong email or password.'); }
  if (N_(u.v[U.blocked])) fail_('This account has been blocked. Please contact support.');
  clearFail_(email);
  return Object.assign({ ok: true, token: newSession_(u) }, state_(u));
}
function list_(d, u) {
  purgeTrash_(u);
  return Object.assign({ ok: true }, state_(u));
}

/* ---------- forgot password: a NEW temporary password is e-mailed. The old password keeps working until the new one is used. ---------- */
function tempPw_() {                                                // 10 easy-to-read characters (no 0/O, 1/l/I)
  const A = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', x = newId_(); let p = '';
  for (let i = 0; i < 10; i++) p += A.charAt(parseInt(x.substr(i * 3, 3), 16) % A.length);
  return p;
}
function forgotPassword_(d) {
  const email = String(d.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail_('Please enter a valid email.');
  const cache = CacheService.getScriptCache(), mk = 'mails_' + sha256_(email).slice(0, 40), n = N_(cache.get(mk));
  if (n >= RESET_MAX_MAILS) fail_('Too many requests. Please try again after 1 hour.');
  cache.put(mk, String(n + 1), 3600);
  const u = findUser_('email', email);
  if (u) {                                                          // same answer whether the account exists or not
    const pw = tempPw_(), salt = newId_();
    cache.put('reset_' + sha256_(email).slice(0, 40), JSON.stringify({ salt, hash: hashPw_(pw, salt) }), RESET_MINUTES * 60);
    try {
      MailApp.sendEmail({
        to: email, name: APP_NAME, subject: APP_NAME + ' - your new password',
        body: 'Hello ' + u.v[U.name] + ',\n\nYour new password is:\n\n    ' + pw + '\n\nLog in with it within ' + RESET_MINUTES +
              ' minutes, then change it from Menu > Account.\nYour old password also keeps working until you use this one.\n\nIf you did not ask for this, you can ignore this email.'
      });
    } catch (e) { console.error('Mail failed: ' + e); fail_('The email could not be sent. Please try again later.'); }
  }
  return { ok: true };
}

/* ---------- account ---------- */
function updateName_(d, u) {
  const n = String(d.name || '').trim().replace(/\s+/g, ' ').slice(0, 60); if (!n) fail_('Please enter a name.');
  u.v[U.name] = n; saveUser_(u); return { ok: true, user: publicUser_(u) };
}
function changePassword_(d, u) {
  checkThrottle_(String(u.v[U.email]));
  if (!pwOk_(u, d.oldPassword)) { addFail_(String(u.v[U.email])); fail_('Current password is incorrect.'); }
  clearFail_(String(u.v[U.email]));
  const np = String(d.newPassword || '');
  if (np.length < 6) fail_('New password must be at least 6 characters.');
  if (np.length > 100) fail_('Password is too long.');
  const salt = newId_(); u.v[U.hash] = hashPw_(np, salt); u.v[U.salt] = salt; saveUser_(u);
  return { ok: true };
}
function deleteAccount_(d, u) {
  if (!pwOk_(u, d.password)) fail_('Password is incorrect.');
  if (isAdmin_(u)) fail_('The admin account cannot be deleted.');
  wipeUser_(u); return { ok: true };
}
function wipeUser_(u) {
  const uid = uid_(u);
  try { DriveApp.getFolderById(u.v[U.driveId]).setTrashed(true); } catch (e) {}
  const us = db_().getSheetByName(uid); if (us) db_().deleteSheet(us);
  const ss = db_().getSheetByName('Shares'), last = ss.getLastRow();
  if (last > 1) { const vals = ss.getRange(2, 1, last - 1, SHEAD.length).getValues().filter(r => String(r[1]) !== uid); ss.getRange(2, 1, last, SHEAD.length).clearContent(); if (vals.length) ss.getRange(2, 1, vals.length, SHEAD.length).setValues(vals); }
  dropSessions_(r => String(r[1]) === uid);
  usersSheet_().deleteRow(u.row);
}

/* ============================ Item rows ============================ */
function userSheet_(uid) { const s = db_().getSheetByName(uid); if (!s) fail_('Account data not found.', 'AUTH'); return s; }
function readRows_(uid) {
  const sh = userSheet_(uid), last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, RHEAD.length).getValues().filter(r => r[R.id] !== '');
}
function writeRows_(uid, rows) {
  const sh = userSheet_(uid), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, RHEAD.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, RHEAD.length).setValues(rows);
}
const idOf_ = s => (String(s).match(/[?&]id=([\w-]+)/) || [])[1] || String(s);
const cleanName_ = s => String(s || 'file').replace(/[\\\/\u0000-\u001f]/g, '').trim().slice(0, 120) || 'file';
function trashFile_(id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
function canonFolder_(u, name) {
  const n = String(name || 'My Drive').toLowerCase();
  if (n === 'favorites' || n === 'trash') return 'My Drive';
  const all = DEFAULT_FOLDERS.concat(customFolders_(u)), f = all.find(x => x.toLowerCase() === n);
  return f || 'My Drive';
}
function purgeTrash_(u) {
  const rows = readRows_(uid_(u)), cut = Date.now() - TRASH_DAYS * 864e5;
  const dead = rows.filter(r => String(r[R.folder]).toLowerCase() === 'trash' && N_(r[R.trashed]) && N_(r[R.trashed]) < cut);
  if (!dead.length) return;
  dead.forEach(r => trashFile_(r[R.id]));
  const ids = new Set(dead.map(r => r[R.id])); writeRows_(uid_(u), rows.filter(r => !ids.has(r[R.id])));
}

/* ============================ Upload ============================ */
function checkCanUpload_(u, add) {
  const rows = readRows_(uid_(u));
  if (rows.length >= MAX_ITEMS_PER_USER) fail_('Upload is not available right now. Please contact support.');
  if (usedOf_(rows) + add > limitBytes_(u)) fail_('Your storage is full. Delete some items and try again.');
}
function addItem_(u, o) {
  const sh = userSheet_(uid_(u));
  const row = [o.id, o.folder, Date.now(), o.size, o.name, o.kind, o.dur || 0, 0, ''];
  sh.appendRow(row);
  return row;
}
function upload_(d, u) {
  const mime = String(d.mime || 'image/jpeg');
  if (!/^image\//.test(mime)) fail_('Only image files can be uploaded here.');
  const bytes = Utilities.base64Decode(String(d.data || '')), size = bytes.length;
  if (!size) fail_('The file is empty.');
  if (size > MAX_IMAGE_BYTES) fail_('This image is too large.');
  checkCanUpload_(u, size);
  const name = cleanName_(d.name), folder = canonFolder_(u, d.folder);
  const file = DriveApp.getFolderById(u.v[U.driveId]).createFile(Utilities.newBlob(bytes, mime, name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  addItem_(u, { id: file.getId(), folder, size, name, kind: 'i' });
  return { ok: true, url: file.getId(), size };
}

/* Video: step 1 - open a resumable upload session in the user's Drive folder */
function vStart_(d, u) {
  const size = N_(d.size), mime = /^video\//.test(String(d.mime)) ? String(d.mime) : 'video/mp4';
  if (size < 1) fail_('The video is empty.');
  if (size > MAX_VIDEO_BYTES) fail_('Video must be 100 MB or smaller.');
  checkCanUpload_(u, size);
  const res = UrlFetchApp.fetch(UPLOAD_HOST + 'uploadType=resumable&fields=id,size', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': mime, 'X-Upload-Content-Length': String(size) },
    payload: JSON.stringify({ name: cleanName_(d.name), parents: [u.v[U.driveId]], mimeType: mime })
  });
  if (res.getResponseCode() !== 200) fail_('Could not start the video upload (' + res.getResponseCode() + ').');
  const h = res.getAllHeaders(), session = h.Location || h.location;
  if (!session) fail_('Could not start the video upload.');
  return { ok: true, session };
}

/* Video: step 2 - relay one chunk (or probe how much Drive already has) */
function vChunk_(d, u) {
  const session = String(d.session || ''), total = N_(d.total), offset = N_(d.offset);
  if (session.indexOf(UPLOAD_HOST) !== 0) fail_('Invalid upload session.');
  if (total < 1 || total > MAX_VIDEO_BYTES) fail_('Video must be 100 MB or smaller.');
  const opts = { method: 'put', muteHttpExceptions: true, followRedirects: false, headers: {} };
  if (d.probe) { opts.headers['Content-Range'] = 'bytes */' + total; opts.payload = ''; }
  else {
    const bytes = Utilities.base64Decode(String(d.data || '')), len = bytes.length;
    if (!len || offset + len > total) fail_('Invalid chunk.');
    opts.payload = bytes; opts.contentType = 'application/octet-stream';
    opts.headers['Content-Range'] = 'bytes ' + offset + '-' + (offset + len - 1) + '/' + total;
  }
  const res = UrlFetchApp.fetch(session, opts), code = res.getResponseCode();
  if (code === 308) {
    const h = res.getAllHeaders(), rg = String(h.Range || h.range || '');
    const m = rg.match(/-(\d+)$/);
    return { ok: true, done: false, next: m ? Number(m[1]) + 1 : 0 };
  }
  if (code === 200 || code === 201) return locked_(() => finishVideo_(u, d, JSON.parse(res.getContentText())));
  if (code === 404 || code === 410) fail_('The upload session expired. Please try again.');
  fail_('Upload failed (' + code + '). Please try again.');
}
function finishVideo_(u, d, meta) {
  const id = meta.id; if (!id) fail_('Upload failed.');
  const rows = readRows_(uid_(u)), have = rows.find(r => r[R.id] === id);
  if (have) return { ok: true, done: true, url: id, size: N_(have[R.size]) };   // retry after a lost response
  const file = DriveApp.getFileById(id), size = file.getSize();
  if (size > MAX_VIDEO_BYTES || usedOf_(rows) + size > limitBytes_(u) || rows.length >= MAX_ITEMS_PER_USER) { file.setTrashed(true); fail_('Your storage is full or the video is too large.'); }
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  addItem_(u, { id, folder: canonFolder_(u, d.folder), size, name: cleanName_(d.name), kind: 'v', dur: Math.round(N_(d.dur)) });
  return { ok: true, done: true, url: id, size };
}

/* ============================ Folders ============================ */
function validFolderName_(n) {
  n = String(n || '').trim().replace(/\s+/g, ' ');
  if (!n) fail_('Please enter a folder name.');
  if (n.length > 40) fail_('Folder name is too long.');
  return n;
}
function createFolder_(d, u) {
  const n = validFolderName_(d.name), all = DEFAULT_FOLDERS.concat(customFolders_(u));
  if (all.some(x => x.toLowerCase() === n.toLowerCase())) fail_('A folder with this name already exists.');
  const a = customFolders_(u); a.push(n); u.v[U.folders] = JSON.stringify(a); saveUser_(u);
  return { ok: true, name: n };
}
function renameFolder_(d, u) {
  const from = String(d.name || ''), a = customFolders_(u), i = a.findIndex(x => x.toLowerCase() === from.toLowerCase());
  if (i < 0) fail_('This folder cannot be renamed.');
  const n = validFolderName_(d.newName);
  if (n.toLowerCase() !== from.toLowerCase() && DEFAULT_FOLDERS.concat(a).some(x => x.toLowerCase() === n.toLowerCase())) fail_('A folder with this name already exists.');
  const old = a[i]; a[i] = n; u.v[U.folders] = JSON.stringify(a);
  const rows = readRows_(uid_(u)); rows.forEach(r => { if (String(r[R.folder]).toLowerCase() === old.toLowerCase()) r[R.folder] = n; }); writeRows_(uid_(u), rows);
  const p = prefsOf_(u);
  if (p.covers && p.covers[old]) { p.covers[n] = p.covers[old]; delete p.covers[old]; }
  if (p.order) p.order = p.order.map(x => x.toLowerCase() === old.toLowerCase() ? n : x);
  setPrefs_(u, p); return { ok: true, name: n };
}
function deleteFolder_(d, u) {
  const from = String(d.name || ''), a = customFolders_(u), i = a.findIndex(x => x.toLowerCase() === from.toLowerCase());
  if (i < 0) fail_('This folder cannot be deleted.');
  const old = a[i]; a.splice(i, 1); u.v[U.folders] = JSON.stringify(a);
  const rows = readRows_(uid_(u));
  rows.forEach(r => {
    if (String(r[R.folder]).toLowerCase() !== old.toLowerCase()) return;
    if (d.deletePhotos) { r[R.folder] = 'Trash'; r[R.trashed] = Date.now(); r[R.fav] = 0; } else r[R.folder] = 'My Drive';
  });
  writeRows_(uid_(u), rows);
  const p = prefsOf_(u); if (p.covers) delete p.covers[old]; if (p.order) p.order = p.order.filter(x => x.toLowerCase() !== old.toLowerCase());
  setPrefs_(u, p); return { ok: true };
}
function setCover_(d, u) {
  const p = prefsOf_(u); p.covers = p.covers || {}; p.covers[String(d.folder)] = idOf_(d.id); setPrefs_(u, p); return { ok: true, prefs: p };
}
function setOrder_(d, u) {
  const p = prefsOf_(u); p.order = (Array.isArray(d.order) ? d.order : []).map(String).slice(0, 200); setPrefs_(u, p); return { ok: true, prefs: p };
}

/* ============================ Item actions ============================ */
const idSet_ = items => new Set((Array.isArray(items) ? items : []).map(x => idOf_(x && x.url !== undefined ? x.url : x)));
function remove_(d, u) {
  const ids = idSet_(d.items), rows = readRows_(uid_(u)), keep = [];
  rows.forEach(r => {
    if (!ids.has(r[R.id])) return keep.push(r);
    if (String(r[R.folder]).toLowerCase() === 'trash') { trashFile_(r[R.id]); return; }      // already in Trash: delete forever
    r[R.folder] = 'Trash'; r[R.trashed] = Date.now(); r[R.fav] = 0; keep.push(r);
  });
  writeRows_(uid_(u), keep); return { ok: true };
}
function move_(d, u) {
  const to = canonFolder_(u, d.to), ids = idSet_(d.items), rows = readRows_(uid_(u));
  if (String(d.to || '').toLowerCase() === 'trash' || String(d.to || '').toLowerCase() === 'favorites') fail_('You cannot move items there.');
  rows.forEach(r => { if (ids.has(r[R.id])) { r[R.folder] = to; r[R.trashed] = ''; } });
  writeRows_(uid_(u), rows); return { ok: true };
}
function favorite_(d, u) {
  const id = idOf_(d.url), rows = readRows_(uid_(u)); let hit = false;
  rows.forEach(r => { if (r[R.id] === id && String(r[R.folder]).toLowerCase() !== 'trash') { r[R.fav] = d.on ? 1 : 0; hit = true; } });
  if (!hit) fail_('Item not found.');
  writeRows_(uid_(u), rows); return { ok: true };
}
function emptyTrash_(d, u) {
  const rows = readRows_(uid_(u)), keep = [];
  rows.forEach(r => { if (String(r[R.folder]).toLowerCase() === 'trash') trashFile_(r[R.id]); else keep.push(r); });
  writeRows_(uid_(u), keep); return { ok: true };
}
function info_(d, u) {
  const id = idOf_(d.id), rows = readRows_(uid_(u)), r = rows.find(x => x[R.id] === id);
  if (!r) fail_('Item not found.');
  if (!N_(r[R.size])) { try { r[R.size] = DriveApp.getFileById(id).getSize(); writeRows_(uid_(u), rows); } catch (e) {} }
  return { ok: true, meta: [N_(r[R.ts]), N_(r[R.size]), r[R.name], r[R.kind] === 'v' ? 'v' : 'i', N_(r[R.dur])] };
}
function recountFor_(u) {
  const rows = readRows_(uid_(u)), t0 = Date.now(); let left = 0;
  rows.forEach(r => {
    if (N_(r[R.size])) return;
    if (Date.now() - t0 > 45000) { left++; return; }
    try { r[R.size] = DriveApp.getFileById(r[R.id]).getSize(); } catch (e) { r[R.size] = 1; }
  });
  writeRows_(uid_(u), rows); u.v[U.used] = usedOf_(rows); saveUser_(u);
  return { ok: true, left };
}
function recount_(d, u) { return recountFor_(u); }

/* ============================ Sharing ============================ */
function share_(d, u) {
  const type = d.type === 'folder' ? 'folder' : 'photo', target = type === 'photo' ? idOf_(d.target) : String(d.target || '');
  if (type === 'photo') { if (!readRows_(uid_(u)).some(r => r[R.id] === target && String(r[R.folder]).toLowerCase() !== 'trash')) fail_('Item not found.'); }
  else if (!DEFAULT_FOLDERS.concat(customFolders_(u)).some(x => x.toLowerCase() === target.toLowerCase())) fail_('Folder not found.');
  const sh = db_().getSheetByName('Shares'), last = sh.getLastRow();
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, SHEAD.length).getValues();
    const f = vals.find(r => String(r[1]) === uid_(u) && r[2] === type && String(r[3]).toLowerCase() === target.toLowerCase() && String(r[5]) === '1');
    if (f) return { ok: true, token: f[0] };
  }
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 28);
  sh.appendRow([token, uid_(u), type, target, Date.now(), 1]); return { ok: true, token };
}
function unshare_(d, u) {
  const sh = db_().getSheetByName('Shares'), last = sh.getLastRow(); if (last < 2) return { ok: true };
  const vals = sh.getRange(2, 1, last - 1, SHEAD.length).getValues();
  vals.forEach((r, i) => { if (String(r[0]) === String(d.token) && String(r[1]) === uid_(u)) sh.getRange(i + 2, 6).setValue(0); });
  return { ok: true };
}
function sharedView_(token) {
  const sh = db_().getSheetByName('Shares'), last = sh.getLastRow(); if (last < 2) fail_('This link is not available.');
  const s = sh.getRange(2, 1, last - 1, SHEAD.length).getValues().find(r => String(r[0]) === token && String(r[5]) === '1');
  if (!s) fail_('This link is no longer available.');
  const u = findUser_('uid', s[1]); if (!u || N_(u.v[U.blocked])) fail_('This link is no longer available.');
  const rows = readRows_(String(s[1])).filter(r => String(r[R.folder]).toLowerCase() !== 'trash'), meta = {}; let ids, name;
  if (s[2] === 'photo') { const r = rows.find(x => x[R.id] === String(s[3])); if (!r) fail_('This link is no longer available.'); ids = [r[R.id]]; name = r[R.name]; }
  else { const t = String(s[3]).toLowerCase(); const m = rows.filter(r => t === 'favorites' ? N_(r[R.fav]) : String(r[R.folder]).toLowerCase() === t); ids = m.map(r => r[R.id]); name = String(s[3]); }
  rows.forEach(r => { if (ids.indexOf(r[R.id]) >= 0) meta[r[R.id]] = [N_(r[R.ts]), N_(r[R.size]), r[R.name], r[R.kind] === 'v' ? 'v' : 'i', N_(r[R.dur])]; });
  return { ok: true, name, ids, meta };
}

/* ============================ Admin ============================ */
function needAdmin_(u) { if (!isAdmin_(u)) fail_('Not allowed.'); }
function adminUsers_(d, u) {
  needAdmin_(u);
  const sh = usersSheet_(), last = sh.getLastRow(); if (last < 2) return { ok: true, users: [] };
  const users = sh.getRange(2, 1, last - 1, UHEAD.length).getValues().map(v => {
    const s = db_().getSheetByName(String(v[U.uid]));
    return { name: v[U.name], email: v[U.email], photos: s ? Math.max(0, s.getLastRow() - 1) : 0, used: N_(v[U.used]), limitGB: N_(v[U.limitGB]) || DEFAULT_LIMIT_GB, verified: !!N_(v[U.verified]), blocked: !!N_(v[U.blocked]) };
  });
  return { ok: true, users };
}
function adminTarget_(d) { const t = findUser_('email', String(d.email || '').toLowerCase()); if (!t) fail_('User not found.'); return t; }
function adminSet_(d, u) {
  needAdmin_(u); const t = adminTarget_(d);
  if (d.limitGB !== undefined) { const g = N_(d.limitGB); if (g <= 0) fail_('Invalid limit.'); t.v[U.limitGB] = g; }
  if (d.blocked !== undefined) { if (isAdmin_(t)) fail_('The admin account cannot be blocked.'); t.v[U.blocked] = d.blocked ? 1 : 0; }
  saveUser_(t); return { ok: true };
}
function adminRecount_(d, u) { needAdmin_(u); return recountFor_(adminTarget_(d)); }
function adminPhotos_(d, u) { needAdmin_(u); const s = state_(adminTarget_(d)); return { ok: true, folders: s.folders }; }
function adminDelete_(d, u) {
  needAdmin_(u); const t = adminTarget_(d), id = idOf_(d.id), rows = readRows_(uid_(t));
  if (!rows.some(r => r[R.id] === id)) fail_('Item not found.');
  trashFile_(id); const keep = rows.filter(r => r[R.id] !== id); writeRows_(uid_(t), keep); t.v[U.used] = usedOf_(keep); saveUser_(t); return { ok: true };
}
