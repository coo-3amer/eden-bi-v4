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
