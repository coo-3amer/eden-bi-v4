/************************************************
 * EDEN BI V4 — CommissionSetup.js
 *
 * Settings of the internal commission calculator (Transactions ›
 * Internal Commissions): sales roles + monthly targets, the rate tiers
 * (In-direct / Direct / Private), and the staff list of every department
 * with how each person is paid. Kept in Script Properties (JSON).
 * The calculation itself runs in the page from the Deals data.
 *   Sales team  : target by role; rate tier by the month's achievement.
 *   Other depts : no target — % of sales, EGP per contract, or EGP per 1M.
 ************************************************/

const COMM_SETUP_KEY_ = 'COMM_SETUP_V1';
const COMM_DEPTS_ = ['Sales', 'Sales Management', 'Operations', 'Customer Service / Contracts', 'Finance / Collection', 'Marketing', 'HR', 'Secretary', 'Buffet', 'Other'];

function commDefaultSetup_() {
  return {
    roles: [{ role: 'Senior', target: 8000000 }, { role: 'Supervisor', target: 10000000 }, { role: 'Team Leader', target: 10000000 }],
    // [minimum achievement (1 = 100%), rate]
    indirect: [[0, 0.006], [0.5001, 0.007], [0.7501, 0.0075], [1, 0.008], [1.5001, 0.009]],
    direct: [[0, 0.015], [0.5001, 0.017], [0.7501, 0.02]],
    privateRate: 0.03,
    shareSplit: 0.5,
    project: 'EDEN WALK',
    staff: [],   // {name, dept, role, scheme: 'tiers'|'pct'|'perDeal'|'perMillion', value, scope: 'all'|'team', active}
    depts: COMM_DEPTS_
  };
}

function commAdmin_(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Internal commissions are for admins only.');
  return user;
}

function commReadSetup_() {
  const def = commDefaultSetup_();
  let s = null;
  try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty(COMM_SETUP_KEY_) || 'null'); } catch (e) {}
  if (!s) return def;
  Object.keys(def).forEach(k => { if (s[k] === undefined) s[k] = def[k]; });
  COMM_DEPTS_.forEach(d => { if (s.depts.indexOf(d) < 0) s.depts.push(d); });
  s.depts = s.depts.filter(d => d !== 'Other').concat(s.depts.indexOf('Other') >= 0 ? ['Other'] : []);   // "Other" always last
  return s;
}

function getCommSetup(authToken) {
  commAdmin_(authToken);
  return commReadSetup_();
}

function saveCommSetup(authToken, setup) {
  const user = commAdmin_(authToken);
  const num = (v, d) => { const n = Number(v); return isFinite(n) ? n : d; };
  const tiers = list => (Array.isArray(list) ? list : []).map(t => [num(t[0], 0), num(t[1], 0)])
    .filter(t => t[1] >= 0 && t[1] < 1).sort((a, b) => a[0] - b[0]);
  const s = setup || {};
  const clean = {
    roles: (s.roles || []).map(r => ({ role: String(r.role || '').trim(), target: Math.max(0, num(r.target, 0)) })).filter(r => r.role),
    indirect: tiers(s.indirect),
    direct: tiers(s.direct),
    privateRate: Math.max(0, Math.min(1, num(s.privateRate, 0.03))),
    shareSplit: Math.max(0, Math.min(1, num(s.shareSplit, 0.5))),
    project: String(s.project || '').trim(),
    depts: (s.depts || COMM_DEPTS_).map(d => String(d || '').trim()).filter(Boolean),
    staff: (s.staff || []).map(p => ({
      name: String(p.name || '').trim(), dept: String(p.dept || 'Sales').trim(), role: String(p.role || '').trim(),
      scheme: ['tiers', 'pct', 'perDeal', 'perMillion'].indexOf(p.scheme) >= 0 ? p.scheme : 'tiers',
      value: Math.max(0, num(p.value, 0)), scope: p.scope === 'team' ? 'team' : 'all', active: p.active !== false
    })).filter(p => p.name)
  };
  if (!clean.indirect.length || !clean.direct.length) throw new Error('Keep at least one rate in In-direct and Direct.');
  const text = JSON.stringify(clean);
  if (text.length > 8500) throw new Error('Too many staff rows to save (' + clean.staff.length + ').');
  PropertiesService.getScriptProperties().setProperty(COMM_SETUP_KEY_, text);
  auditLog_(user, 'Commission setup saved', { details: clean.staff.length + ' staff · ' + clean.roles.map(r => r.role + ' ' + r.target).join(', ') });
  return clean;
}

/* ---------- Manual eligibility per deal ----------
   Besides the automatic rule (DP paid % range, not cancelled), an admin can tick a
   deal in (counts) or out (doesn't count). Kept in the hidden sheet "_Comm Overrides":
   Key | Include | Label | By | At.  Key = project|unit|client|reservation date. */
const COMM_OV_SHEET_ = '_Comm Overrides';
function commOvSheet_() {
  const ss = targetsSpreadsheet_();
  let sh = ss.getSheetByName(COMM_OV_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(COMM_OV_SHEET_);
    sh.getRange(1, 1, 1, 5).setValues([['Key', 'Include', 'Deal', 'By', 'At']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}
function getCommOverrides(authToken) {
  commAdmin_(authToken);
  const sh = commOvSheet_(), last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 5).getValues().forEach(r => {
    const k = String(r[0] || '').trim();
    if (k) out[k] = { inc: r[1] === true || String(r[1]).toUpperCase() === 'TRUE', by: String(r[3] || ''), at: r[4] instanceof Date ? r[4].toISOString() : String(r[4] || '') };
  });
  return out;
}
/* items: [{key, inc: true|false|null, label}] — null removes the override (back to automatic). */
function setCommOverrides(authToken, items) {
  const user = commAdmin_(authToken);
  items = (Array.isArray(items) ? items : []).map(i => ({ key: String(i && i.key || '').trim(), inc: i && i.inc === null ? null : !!(i && i.inc), label: String(i && i.label || '').slice(0, 200) })).filter(i => i.key);
  if (!items.length) return getCommOverrides(authToken);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = commOvSheet_(), last = sh.getLastRow();
    let rows = last >= 2 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
    const now = new Date(), who = (user && (user.name || user.username)) || '';
    let inc = 0, exc = 0, auto = 0;
    items.forEach(it => {
      const idx = rows.findIndex(r => String(r[0]).trim() === it.key);
      if (it.inc === null) { if (idx >= 0) rows.splice(idx, 1); auto++; return; }
      const rec = [it.key, it.inc, it.label, who, now];
      if (idx >= 0) rows[idx] = rec; else rows.push(rec);
      it.inc ? inc++ : exc++;
    });
    if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
    const extra = Math.max(0, (last - 1) - rows.length);
    if (extra) sh.getRange(2 + rows.length, 1, extra, 5).clearContent();
    auditLog_(user, 'Commission deals chosen', { details: (inc ? inc + ' included ' : '') + (exc ? exc + ' excluded ' : '') + (auto ? auto + ' back to automatic' : '') + (items.length === 1 ? ' · ' + items[0].label : '') });
  } finally { lock.releaseLock(); }
  return getCommOverrides(authToken);
}

/* ---------- Sales / Sales Manager: their own commission view ----------
   Only what is needed to estimate their own commission: rates, their own
   staff entry, the manual choices on their own deals, and their closed months. */
function getMyCommSetup(authToken) {
  const user = validateAuthToken_(authToken);
  if (!isSalesScopedUser_(user)) throw new Error('For sales accounts only.');
  const me = salesNameKey_(user.name);
  const s = commReadSetup_();
  const staff = (s.staff || []).filter(p => salesNameKey_(p.name) === me && p.active !== false);
  const setup = { roles: s.roles, indirect: s.indirect, direct: s.direct, privateRate: s.privateRate, shareSplit: s.shareSplit, project: s.project, depts: s.depts, staff: staff };
  // manual choices only for the deals this user can see
  let ov = {};
  try {
    let data = dashCacheRead_('ALL', 6 * 3600 * 1000) || buildDashboardData_({ role: 'Admin' });
    const mine = scopeDashboardForSales_(data, user).rows || [];
    const key = r => [r.project, r.unitCode, r.clientName, (Number(r.year) && Number(r.month)) ? r.year + '-' + String(r.month).padStart(2, '0') + '-' + String(Number(r.day) || 1).padStart(2, '0') : '']
      .map(v => String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' ')).join('|');
    const keys = {}; mine.forEach(r => keys[key(r)] = 1);
    const all = getCommOverridesAll_();
    Object.keys(all).forEach(k => { if (keys[k]) ov[k] = { inc: all[k].inc }; });
  } catch (e) { console.warn('My overrides: ' + e); }
  // closed months for this person
  let closed = [];
  try {
    closed = commCloseRows_().filter(r => salesNameKey_(r[1]) === me).map(r => ({ month: String(r[0]), comm: Number(r[7]) || 0, credited: Number(r[5]) || 0, ach: Number(r[6]) || 0,
      paid: r[12] === true || String(r[12]).toUpperCase() === 'TRUE', paidAt: r[13] instanceof Date ? r[13].toISOString() : String(r[13] || '') }))
      .sort((a, b) => b.month.localeCompare(a.month)).slice(0, 12);
  } catch (e) {}
  return { setup: setup, ov: ov, closed: closed, name: user.name };
}
function getCommOverridesAll_() {
  const sh = commOvSheet_(), last = sh.getLastRow(), out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => { const k = String(r[0] || '').trim(); if (k) out[k] = { inc: r[1] === true || String(r[1]).toUpperCase() === 'TRUE' }; });
  return out;
}

/* ---------- Mark commission paid on deals (Deals sheet › Internal comm) ----------
   req = {items:[{row, unit, set:{sales:true, managers:false, staff:true}}]}
   (old form {items:[{row,unit}], cols:[...], value} still works).
   Checks each row still holds that unit, then writes all the cells in a few calls. */
function setDealCommPaid(authToken, req) {
  const user = commAdmin_(authToken);
  req = req || {};
  const KEYS = ['sales', 'managers', 'staff'];
  let items = (Array.isArray(req.items) ? req.items : []).filter(i => Number(i && i.row) >= DATA_START_ROW);
  if (Array.isArray(req.cols)) {
    const v = req.value !== false;
    items = items.map(i => { const set = {}; req.cols.forEach(c => { c = String(c).toLowerCase(); if (KEYS.indexOf(c) >= 0) set[c] = v; }); return { row: i.row, unit: i.unit, set: set }; });
  }
  items = items.filter(i => i.set && Object.keys(i.set).some(k => KEYS.indexOf(k) >= 0));
  if (!items.length) throw new Error('Pick the deals and which commission (Sales / Managers / Staff).');
  if (items.length > 500) throw new Error('Too many deals at once (max 500).');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  const done = [], skipped = [], counts = {};
  try {
    const sh = getMainDealsSheet_();
    const cache = CacheService.getScriptCache();
    let C = null; try { C = JSON.parse(cache.get('COMM_PAID_COLS') || 'null'); } catch (e) {}
    if (!C) {
      const cols = mainDealsColumns_(sh), at = n => mainDealsCol_(cols, n);
      C = { sales: at(['internal comm|sales']), managers: at(['internal comm|managers']), staff: at(['internal comm|staff']), unit: at(['Unit Code']) };
      cache.put('COMM_PAID_COLS', JSON.stringify(C), 600);
    }
    KEYS.forEach(k => { if (!C[k]) throw new Error('Column "Internal comm › ' + k + '" was not found in Deals.'); });
    const last = sh.getLastRow();
    const norm = v => String(v || '').trim().toUpperCase().replace(/\.+$/, '');
    // units of the rows involved (one read for many rows, a few small reads for a handful)
    const rows = items.map(i => Number(i.row)).filter(r => r <= last);
    const unitAt = {};
    if (rows.length > 15) {
      const lo = Math.min.apply(null, rows), hi = Math.max.apply(null, rows);
      sh.getRange(lo, C.unit, hi - lo + 1, 1).getDisplayValues().forEach((r, k) => unitAt[lo + k] = norm(r[0]));
    } else if (rows.length) {
      sh.getRangeList(rows.map(r => sh.getRange(r, C.unit).getA1Notation())).getRanges().forEach((rg, k) => unitAt[rows[k]] = norm(rg.getDisplayValue()));
    }
    const groups = {};   // "sales|true" -> [A1, …]
    items.forEach(i => {
      const row = Number(i.row), u = norm(i.unit);
      if (row > last || unitAt[row] !== u) { skipped.push(i.unit || ('row ' + row)); return; }
      Object.keys(i.set).forEach(k => { if (KEYS.indexOf(k) < 0) return; const g = k + '|' + (i.set[k] ? '1' : '0'); (groups[g] = groups[g] || []).push(sh.getRange(row, C[k]).getA1Notation()); counts[g] = (counts[g] || 0) + 1; });
      done.push(u);
    });
    Object.keys(groups).forEach(g => {
      const list = sh.getRangeList(groups[g]);
      list.insertCheckboxes();
      list.setValue(g.slice(-1) === '1');
    });
  } finally { lock.releaseLock(); }
  try { clearDashboardCache_(); } catch (e) {}
  const what = Object.keys(counts).map(g => (g.slice(-1) === '1' ? '✓ ' : '✗ ') + g.split('|')[0] + ' ×' + counts[g]).join(', ');
  auditLog_(user, 'Commission paid ticks in Deals', {
    details: what + ' · ' + done.length + ' deal' + (done.length === 1 ? '' : 's') + (done.length <= 8 ? ': ' + done.join(', ') : '') + (skipped.length ? ' · skipped (row changed): ' + skipped.join(', ') : '') });
  return { done: done.length, skipped: skipped };
}
