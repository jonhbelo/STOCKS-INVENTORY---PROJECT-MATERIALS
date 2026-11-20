/**
 * RPI Project Materials Inventory v1.3 - Google Apps Script web app.
 *
 * The page (HTML file "index" in this project) is served by doGet and calls
 * api() through google.script.run. Deploy as: Execute as "Me", Who has access
 * "Anyone within riteproducts.com.ph". The SERVER still checks every request:
 *  - the Google account opening the app must be in the company domain
 *  - logins are verified here; password hashes never leave the server
 *  - every data call needs a session token issued at login (6 hours)
 *  - roles: admin (everything), editor (add records, import, request edits),
 *    viewer (read only); adding users, changing roles and the database are admin only
 *  - 5 wrong passwords lock that account for 15 minutes
 *  - the owner (ADMIN_EMAIL) signs in with their Google account, no password;
 *    the account is created automatically the first time they open the app and
 *    always stays admin
 *  - there is no sign-up: only admins add user accounts
 *
 * Projects: SYSTEMS PROJECTS and HVAC PROJECTS each have their own Google Sheet
 * (see PROJECTS), so items, stock, transactions and edit requests never mix.
 * Accounts and roles are shared: one login works for both.
 *
 * Data: items, received and delivered records go in the "App Data" tab and are
 * copied to the readable Items, Received and Delivered tabs (newest first).
 * Accounts and edit requests are kept in Script Properties; a readable copy of
 * the accounts (name, username, email, role; never passwords) goes to the "Users" tab.
 */

/** Each project has its own Google Sheet, so its items, stock and transactions are kept apart. */
const PROJECTS = {
  systems: { name: 'SYSTEMS PROJECTS', sheet: '1uCAk_P-Y5p7PMT86cIIFjWqItAIgYtv_5fYoYJ1Q5ho' },
  hvac: { name: 'HVAC PROJECTS', sheet: '16OCgM02Q6eiIccptmCUe3MhY8QE8d8WNrE4lpYc5LH0' }
};
function projOf_(p) {
  const k = String(p || 'systems');
  if (!Object.prototype.hasOwnProperty.call(PROJECTS, k)) throw new Error('bad_request');
  return k;
}
function openProj_(p) { return SpreadsheetApp.openById(PROJECTS[projOf_(p)].sheet); }
/** Edit requests made before projects existed belong to Systems. */
function reqProj_(r) { return (r && r.pj) || 'systems'; }
const ADMIN_EMAIL = 'jonh.belo@riteproducts.com.ph';
const ALLOWED_DOMAIN = '@riteproducts.com.ph';
const DATA_TAB = 'App Data';
const SHEET_COLS = ['items', 'received', 'delivered'];
const PROP_COLS = ['reqs'];
const ID_RX = /^[A-Za-z0-9_.@+-]{1,120}$/;
const UN_RX = /^[a-z0-9._-]{3,24}$/;
const HASH_RX = /^[0-9a-f]{64}$/;
const SALT_RX = /^[0-9a-f]{32}$/;
const MAX_FAILS = 5;
const LOCK_SECONDS = 900;
const SESSION_SECONDS = 21600;
const ROLES = ['admin', 'editor', 'viewer'];
const USERS_TAB = 'Users';

/* ---------- one-time setup (run from the editor) ---------- */

/** Creates the App Data tab, clears the example rows, and makes the server secret. */
function setup() {
  Object.keys(PROJECTS).forEach(function (p) {
    const ss = openProj_(p);
    getDataSheet_(ss);
    mirror_(ss, docsOf_(readSheetEntries_(ss)));
    writeUsersTab_(ss);
  });
  pepper_();
}

/* ---------- HTTP entry points ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('RPI Project Materials Inventory')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called by the page through google.script.run. Takes a JSON string, returns {ok} or {err}. */
function api(body) {
  try {
    const who = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    if (who.slice(-ALLOWED_DOMAIN.length) !== ALLOWED_DOMAIN) throw new Error('forbidden');
    const req = JSON.parse(String(body || '{}'));
    return { ok: route_(req, who) };
  } catch (err) {
    const m = String((err && err.message) || err);
    const code = /^[a-z_]{3,30}$/.test(m) ? m : 'server_error';
    if (code === 'server_error') console.error(m);
    return { err: code };
  }
}

function route_(r, who) {
  switch (r.a) {
    case 'status': return status_(r.t, who, !!r.auto);
    case 'salt': return { salt: saltFor_(r.id) };
    case 'login': return login_(r.id, r.hash);
  }
  const me = session_(r.t);
  switch (r.a) {
    case 'logout': CacheService.getScriptCache().remove('t:' + r.t); return true;
    case 'snapshot': return snapshot_(r.p);
    case 'batch': return batch_(r.ops, me, r.p);
    case 'profile': return profile_(me, r.name);
    case 'password': return password_(me, r.oldHash, r.salt, r.hash);
    case 'adduser': needAdmin_(me); return addUser_(r.user);
    case 'role': needAdmin_(me); return setRole_(me, r.id, r.role);
    case 'db': needAdmin_(me); return dbInfo_(r.p);
    case 'rebuild': needAdmin_(me); rebuild_(r.p); return dbInfo_(r.p);
  }
  throw new Error('bad_request');
}

/* ---------- accounts and sessions ---------- */

function props_() { return PropertiesService.getScriptProperties(); }

function getUsers_() {
  const all = props_().getProperties(), out = [];
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('users/') === 0) out.push(JSON.parse(all[k]));
  });
  return out;
}

function putUser_(u) { props_().setProperty('users/' + u.id, JSON.stringify(u)); }

/** admin, editor or viewer. Accounts made before roles existed: admin flag, else editor. */
function roleOf_(u) {
  if (u && u.email === ADMIN_EMAIL) return 'admin';
  return u && ROLES.indexOf(u.role) >= 0 ? u.role : (u && u.admin ? 'admin' : 'editor');
}
function isAdmin_(u) { return roleOf_(u) === 'admin'; }

function pub_(u) {
  const role = roleOf_(u);
  return { id: u.id, email: u.email, username: u.username, name: u.name, role: role,
    admin: role === 'admin', owner: u.email === ADMIN_EMAIL, hasPw: !!u.hash };
}

function findUser_(ident) {
  const k = String(ident || '').trim().toLowerCase();
  if (!k) return null;
  return getUsers_().filter(function (u) { return u.email === k || u.username === k; })[0] || null;
}

function eq_(a, b) {
  a = String(a); b = String(b);
  let d = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return d === 0;
}

function pepper_() {
  let p = props_().getProperty('PEPPER');
  if (!p) {
    p = Utilities.getUuid() + Utilities.getUuid();
    props_().setProperty('PEPPER', p);
  }
  return p;
}

/** The salt for an account. Unknown names get a stable fake salt so the answer does not reveal who has an account. */
function saltFor_(ident) {
  const u = findUser_(ident);
  if (u && SALT_RX.test(String(u.salt))) return u.salt;
  const sig = Utilities.computeHmacSha256Signature(String(ident || '').toLowerCase(), pepper_());
  return sig.map(function (b) { return ('0' + ((b < 0 ? b + 256 : b)).toString(16)).slice(-2); }).join('').slice(0, 32);
}

function failCount_(key) { return Number(CacheService.getScriptCache().get('f:' + key) || 0); }
function addFail_(key) { CacheService.getScriptCache().put('f:' + key, String(failCount_(key) + 1), LOCK_SECONDS); }
function clearFails_(key) { CacheService.getScriptCache().remove('f:' + key); }

function startSession_(u) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  CacheService.getScriptCache().put('t:' + token, JSON.stringify({ uid: u.id }), SESSION_SECONDS);
  return { token: token, user: pub_(u) };
}

function session_(token) {
  const raw = token ? CacheService.getScriptCache().get('t:' + token) : null;
  if (!raw) throw new Error('auth');
  const u = getUsers_().filter(function (x) { return x.id === JSON.parse(raw).uid; })[0];
  if (!u) throw new Error('auth');
  return u;
}

function needAdmin_(me) { if (!isAdmin_(me)) throw new Error('forbidden'); }

/**
 * Who is signed in. When the Google account opening the app is ADMIN_EMAIL and
 * auto is set, the admin is signed in from that Google account (no password).
 */
function status_(token, who, auto) {
  const gAdmin = who === ADMIN_EMAIL;
  let user = null, tok = null;
  try { user = pub_(session_(token)); } catch (e) { /* not signed in */ }
  if (!user && gAdmin && auto) {
    const s = startSession_(ensureAdmin_());
    user = s.user; tok = s.token;
  }
  return { has: gAdmin || getUsers_().length > 0, user: user, token: tok, gAdmin: gAdmin };
}

/** The admin account for ADMIN_EMAIL, created the first time it is needed. */
function ensureAdmin_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = getUsers_();
    let u = all.filter(function (x) { return x.email === ADMIN_EMAIL; })[0];
    if (u) {
      if (!u.admin || u.role !== 'admin') { u.admin = true; u.role = 'admin'; putUser_(u); }
      return u;
    }
    let base = ADMIN_EMAIL.split('@')[0].toLowerCase().replace(/[^a-z0-9._-]/g, '').slice(0, 20);
    if (base.length < 3) base = 'admin';
    let un = base, n = 1;
    while (all.some(function (x) { return x.username === un; })) un = base + (n++);
    u = { id: Utilities.getUuid().replace(/-/g, '').slice(0, 12), email: ADMIN_EMAIL, username: un,
      name: '', salt: '', hash: '', admin: true, role: 'admin' };
    putUser_(u);
    props_().setProperty('meta/admin', JSON.stringify({ email: ADMIN_EMAIL }));
    syncUsersTab_();
    return u;
  } finally {
    lock.releaseLock();
  }
}

function login_(ident, hash) {
  const key = String(ident || '').trim().toLowerCase().slice(0, 120);
  if (failCount_(key) >= MAX_FAILS) throw new Error('locked');
  const u = findUser_(key);
  if (!u || !HASH_RX.test(String(hash)) || !eq_(u.hash, hash)) {
    addFail_(key);
    throw new Error('bad_login');
  }
  clearFails_(key);
  return startSession_(u);
}

function cleanUser_(u) {
  u = u || {};
  const role = ROLES.indexOf(u.role) >= 0 ? u.role : 'editor';
  const email = String(u.email || '').trim().toLowerCase();
  const username = String(u.username || '').trim().toLowerCase();
  const name = String(u.name || '').trim().slice(0, 80);
  if (!/^\S+@\S+\.\S+$/.test(email) || email.slice(-ALLOWED_DOMAIN.length) !== ALLOWED_DOMAIN) throw new Error('bad_user');
  if (!UN_RX.test(username) || !SALT_RX.test(String(u.salt)) || !HASH_RX.test(String(u.hash))) throw new Error('bad_user');
  return {
    id: Utilities.getUuid().replace(/-/g, '').slice(0, 12),
    email: email, username: username, name: name, salt: u.salt, hash: u.hash,
    role: role, admin: role === 'admin'
  };
}

function addUser_(user) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const nu = cleanUser_(user), all = getUsers_();
    if (nu.email === ADMIN_EMAIL || all.some(function (x) { return x.email === nu.email; })) throw new Error('email_taken');
    if (all.some(function (x) { return x.username === nu.username; })) throw new Error('username_taken');
    putUser_(nu);
    syncUsersTab_();
    return { id: nu.id, role: nu.role };
  } finally {
    lock.releaseLock();
  }
}

/** Admin changes another account's role. The owner account always stays admin. */
function setRole_(me, id, role) {
  if (ROLES.indexOf(role) < 0) throw new Error('bad_request');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const u = getUsers_().filter(function (x) { return x.id === id; })[0];
    if (!u) throw new Error('not_found');
    if (u.email === ADMIN_EMAIL || u.id === me.id) throw new Error('forbidden');
    u.role = role; u.admin = role === 'admin';
    putUser_(u);
    syncUsersTab_();
    return pub_(u);
  } finally {
    lock.releaseLock();
  }
}

function profile_(me, name) {
  me.name = String(name || '').trim().slice(0, 80);
  putUser_(me);
  syncUsersTab_();
  return true;
}

/** Writes the readable Users tab in every project's sheet. */
function syncUsersTab_() {
  Object.keys(PROJECTS).forEach(function (p) { writeUsersTab_(openProj_(p)); });
}

/** Writes the readable Users tab: name, username, email, role. Never salts or hashes. */
function writeUsersTab_(ss) {
  let sh = ss.getSheetByName(USERS_TAB);
  if (!sh) {
    sh = ss.insertSheet(USERS_TAB);
    sh.getRange(1, 1, 1, 4).setValues([['Name', 'Username', 'Email', 'Role']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const order = { admin: 0, editor: 1, viewer: 2 };
  const rows = getUsers_().map(pub_).sort(function (a, b) {
    return order[a.role] - order[b.role] || String(a.name || a.email).localeCompare(String(b.name || b.email));
  }).map(function (u) {
    return [u.name || '', u.username || '', u.email, u.role.charAt(0).toUpperCase() + u.role.slice(1)];
  });
  writeRows_(sh, rows, 4);
}

function password_(me, oldHash, salt, hash) {
  const key = 'p:' + me.id;
  if (failCount_(key) >= MAX_FAILS) throw new Error('locked');
  if (!HASH_RX.test(String(oldHash)) || !eq_(me.hash, oldHash)) { addFail_(key); throw new Error('bad_login'); }
  if (!SALT_RX.test(String(salt)) || !HASH_RX.test(String(hash))) throw new Error('bad_user');
  clearFails_(key);
  me.salt = salt; me.hash = hash;
  putUser_(me);
  return true;
}

/* ---------- data ---------- */

function snapshot_(proj) {
  proj = projOf_(proj);
  const ss = openProj_(proj), s = readSheetEntries_(ss), out = {};
  SHEET_COLS.forEach(function (c) { out[c] = s[c].map(function (e) { return e.data; }); });
  const all = props_().getProperties();
  out.reqs = [];
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('reqs/') !== 0) return;
    const r = JSON.parse(all[k]);
    if (reqProj_(r) === proj) out.reqs.push(r);
  });
  out.users = getUsers_().map(pub_);
  return out;
}

function parse_(path) {
  const s = String(path), i = s.indexOf('/');
  const col = s.slice(0, i), id = s.slice(i + 1);
  if (i < 1 || (SHEET_COLS.indexOf(col) < 0 && PROP_COLS.indexOf(col) < 0) || !ID_RX.test(id)) {
    throw new Error('invalid_argument');
  }
  return { col: col, id: id };
}

/* Field types per collection. 's' = text (max 200), 'n' = finite number, 'id' = record-id text (ID_RX). */
const FIELDS = {
  items: { id: 'id', ts: 'n', name: 's', unit: 's', begin: 'n', min: 'n' },
  received: { id: 'id', ts: 'n', tx: 'id', date: 's', a: 's', b: 's', inv: 's', po: 's', jo: 's', ref: 's', itemId: 's', qty: 'n' },
  delivered: { id: 'id', ts: 'n', tx: 'id', date: 's', a: 's', jo: 's', ref: 's', itemId: 's', qty: 'n' },
  reqs: { id: 'id', by: 's', key: 's', kind: 's', tid: 'id', label: 's', st: 's', at: 'n', pj: 's' }
};
const REQ_STATES = ['pending', 'approved', 'denied', 'done'];
const REQ_KEEP_MS = 30 * 24 * 3600 * 1000;

/** Rejects records the page could not render safely (wrong id, odd types, nested objects). */
function checkRecord_(col, id, d) {
  if (!d || typeof d !== 'object' || Array.isArray(d)) throw new Error('invalid_argument');
  if (d.id !== id) throw new Error('invalid_argument');
  const spec = FIELDS[col];
  Object.keys(d).forEach(function (k) {
    const v = d[k], t = spec[k];
    if (v === null || v === undefined) return;
    if (t === 'n') { if (typeof v !== 'number' || !isFinite(v)) throw new Error('invalid_argument'); }
    else if (t === 'id') { if (typeof v !== 'string' || !ID_RX.test(v)) throw new Error('invalid_argument'); }
    else if (typeof v === 'string') { if (v.length > 200) throw new Error('invalid_argument'); }
    else if (t === 's' || (typeof v !== 'number' && typeof v !== 'boolean')) throw new Error('invalid_argument');
    else if (typeof v === 'number' && !isFinite(v)) throw new Error('invalid_argument');
  });
  if (col === 'reqs' && REQ_STATES.indexOf(d.st) < 0) throw new Error('invalid_argument');
}

/** Edit requests: a non-admin may only file their own pending request, or mark their own approved one done. */
function checkReq_(me, existing, next) {
  if (isAdmin_(me)) return;
  if (next === undefined) throw new Error('forbidden');
  if (!existing) {
    if (next.by !== me.id || next.st !== 'pending') throw new Error('forbidden');
    return;
  }
  if (existing.by !== me.id || next.by !== me.id || next.key !== existing.key ||
      existing.st !== 'approved' || next.st !== 'done') throw new Error('forbidden');
}

/** Edit-request key for a record: items by item id, received/delivered by transaction id. */
function recKey_(col, id, d) {
  return col + ':' + (col === 'items' ? id : ((d && d.tx) || id));
}

/** ops: [{t:'set'|'create'|'update'|'delete', p:'collection/id', d:{...}}]. Returns one result per op. */
function batch_(ops, me, proj) {
  if (!Array.isArray(ops) || ops.length > 300) throw new Error('bad_request');
  if (roleOf_(me) === 'viewer') throw new Error('forbidden');
  proj = projOf_(proj);
  const admin = isAdmin_(me);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const props = props_(), pcur = props.getProperties(), pset = {}, pdel = {};
    const ss = openProj_(proj);
    const entries = readSheetEntries_(ss), idx = {};
    SHEET_COLS.forEach(function (c) {
      idx[c] = {};
      entries[c].forEach(function (e) { idx[c][e.id] = e; });
    });
    // Non-admins may add new records, but changing an existing item or transaction
    // needs an approved edit request for it (same key the page uses: "kind:id").
    const preKeys = {}, approved = {};
    SHEET_COLS.forEach(function (c) {
      entries[c].forEach(function (e) { preKeys[recKey_(c, e.id, e.data)] = 1; });
    });
    if (!admin) {
      Object.keys(pcur).forEach(function (k) {
        if (k.indexOf('reqs/') !== 0) return;
        try {
          const r = JSON.parse(pcur[k]);
          if (r.by === me.id && r.st === 'approved' && reqProj_(r) === proj) approved[r.key] = 1;
        } catch (e) { /* ignore a damaged entry */ }
      });
    }
    let sheetDirty = false;
    const res = [];

    ops.forEach(function (op) {
      try {
        const p = parse_(op.p), isProp = PROP_COLS.indexOf(p.col) >= 0, key = p.col + '/' + p.id;
        let existing = null;
        if (isProp) { if (key in pcur) existing = JSON.parse(pcur[key]); }
        else if (idx[p.col][p.id]) existing = idx[p.col][p.id].data;

        let next;
        if (op.t === 'delete') next = undefined;
        else if (op.t === 'create') {
          if (existing) { res.push({ err: 'already_exists' }); return; }
          next = op.d;
        } else if (op.t === 'update') {
          if (!existing) { res.push({ err: 'invalid_argument' }); return; }
          next = Object.assign({}, existing, op.d);
        } else if (op.t === 'set') next = op.d;
        else throw new Error('bad_op');

        if (next !== undefined) checkRecord_(p.col, p.id, next);
        if (p.col === 'reqs') {
          // A request belongs to the project it was filed in and can only be handled there.
          if ((existing && reqProj_(existing) !== proj) || (next && reqProj_(next) !== proj)) throw new Error('invalid_argument');
          checkReq_(me, existing, next);
        }
        if (!isProp && !admin) {
          [existing, next].forEach(function (d) {
            if (!d) return;
            const k = recKey_(p.col, p.id, d);
            if ((d === existing || preKeys[k]) && !approved[k]) throw new Error('forbidden');
          });
        }

        if (isProp) {
          if (next === undefined) { delete pcur[key]; delete pset[key]; pdel[key] = 1; }
          else {
            const s = JSON.stringify(next);
            if (s.length > 8500) throw new Error('too_large');
            pcur[key] = s; pset[key] = s; delete pdel[key];
          }
        } else {
          if (next === undefined) delete idx[p.col][p.id];
          else idx[p.col][p.id] = { id: p.id, data: next };
          sheetDirty = true;
        }
        res.push({});
      } catch (e) {
        res.push({ err: String((e && e.message) || e) });
      }
    });

    // Drop finished edit requests older than 30 days so Script Properties stays under its size limit.
    const cutoff = Date.now() - REQ_KEEP_MS;
    Object.keys(pcur).forEach(function (k) {
      if (k.indexOf('reqs/') !== 0 || pset[k]) return;
      try {
        const r = JSON.parse(pcur[k]);
        if ((r.st === 'done' || r.st === 'denied') && (r.at || 0) < cutoff) pdel[k] = 1;
      } catch (e) { /* leave a damaged entry alone */ }
    });

    if (Object.keys(pset).length) props.setProperties(pset, false);
    Object.keys(pdel).forEach(function (k) { props.deleteProperty(k); });

    if (sheetDirty) {
      const rows = [], out = { items: [], received: [], delivered: [] }, stamp = new Date().toISOString();
      SHEET_COLS.forEach(function (c) {
        Object.keys(idx[c]).forEach(function (id) {
          out[c].push(idx[c][id].data);
          rows.push([c, id, JSON.stringify(idx[c][id].data), stamp]);
        });
      });
      writeRows_(getDataSheet_(ss), rows, 4);
      mirror_(ss, out);
    }
    return res;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- database tab (admin) ---------- */

/** Connection check and a summary of what is stored where. */
function dbInfo_(proj) {
  proj = projOf_(proj);
  const t0 = Date.now();
  const ss = openProj_(proj);
  const e = readSheetEntries_(ss), sh = ss.getSheetByName(DATA_TAB);
  let last = 0;
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 4, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      const v = r[0] instanceof Date ? r[0].getTime() : Date.parse(String(r[0] || ''));
      if (v > last) last = v;
    });
  }
  const all = props_().getProperties();
  let bytes = 0, users = 0, reqs = 0;
  Object.keys(all).forEach(function (k) {
    bytes += k.length + String(all[k]).length;
    if (k.indexOf('users/') === 0) users++;
    if (k.indexOf('reqs/') === 0) {
      try { if (reqProj_(JSON.parse(all[k])) === proj) reqs++; } catch (e) { /* skip */ }
    }
  });
  return {
    project: PROJECTS[proj].name, name: ss.getName(), url: ss.getUrl(), dataTab: !!sh,
    tabs: ss.getSheets().map(function (s) { return { name: s.getName(), rows: Math.max(0, s.getLastRow() - 1) }; }),
    counts: { items: e.items.length, received: e.received.length, delivered: e.delivered.length, users: users, reqs: reqs },
    updated: last, propsBytes: bytes, ms: Date.now() - t0, at: Date.now()
  };
}

/** Rewrites the readable Items, Received and Delivered tabs from App Data (creates App Data if missing). */
function rebuild_(proj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ss = openProj_(proj);
    getDataSheet_(ss);
    mirror_(ss, docsOf_(readSheetEntries_(ss)));
    writeUsersTab_(ss);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- sheet storage and readable tabs ---------- */

function getDataSheet_(ss) {
  let sh = ss.getSheetByName(DATA_TAB);
  if (!sh) {
    sh = ss.insertSheet(DATA_TAB);
    sh.getRange(1, 1, 1, 4).setValues([['Collection', 'Record ID', 'Record JSON', 'Updated']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readSheetEntries_(ss) {
  const out = { items: [], received: [], delivered: [] }, sh = ss.getSheetByName(DATA_TAB);
  if (!sh) return out;
  const n = sh.getLastRow();
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 3).getValues().forEach(function (r) {
    if (out[r[0]]) {
      try { out[r[0]].push({ id: String(r[1]), data: JSON.parse(r[2]) }); } catch (e) { /* skip a damaged row */ }
    }
  });
  return out;
}

function docsOf_(entries) {
  const o = {};
  SHEET_COLS.forEach(function (c) { o[c] = entries[c].map(function (e) { return e.data; }); });
  return o;
}

/** Replace everything below the header row with rows (n columns wide). */
function writeRows_(sh, rows, n) {
  const max = sh.getMaxRows();
  if (max > 1) sh.getRange(2, 1, max - 1, n).clearContent();
  if (!rows.length) return;
  if (rows.length + 1 > max) sh.insertRowsAfter(max, rows.length + 1 - max);
  sh.getRange(2, 1, rows.length, n).setValues(rows);
}

/* ---------- readable tabs: Items, Received, Delivered (newest first) ---------- */

function dateOf_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : String(s || '');
}

function stampOf_(n) { return n ? new Date(n) : ''; }

/** Newest transaction first; lines inside a transaction keep the order they were entered. */
function byNewest_(list) {
  const top = {};
  list.forEach(function (r) {
    const k = r.tx || r.id;
    top[k] = Math.max(top[k] || 0, r.ts || 0);
  });
  return list.slice().sort(function (a, b) {
    const ta = top[a.tx || a.id], tb = top[b.tx || b.id];
    return tb - ta || (a.ts || 0) - (b.ts || 0);
  });
}

function mirror_(ss, docs) {
  const items = docs.items.slice().sort(function (a, b) { return (b.ts || 0) - (a.ts || 0); });
  const name = {};
  items.forEach(function (i) { name[i.id] = i.name; });

  const iRows = items.map(function (i) {
    return [i.name, i.unit, i.begin || 0, i.min || 0, stampOf_(i.ts)];
  });
  const rRows = byNewest_(docs.received).map(function (r) {
    return [dateOf_(r.date), r.a || '', r.b || '', r.inv || '', r.po || '', r.jo || '',
      name[r.itemId] || '', r.qty, stampOf_(r.ts)];
  });
  const dRows = byNewest_(docs.delivered).map(function (r) {
    return [dateOf_(r.date), r.a || '', r.jo || '', r.ref || '',
      name[r.itemId] || '', r.qty, stampOf_(r.ts)];
  });

  [['Items', iRows, 5], ['Received', rRows, 9], ['Delivered', dRows, 7]].forEach(function (t) {
    const sh = ss.getSheetByName(t[0]);
    if (sh) writeRows_(sh, t[1], t[2]);
  });
}
