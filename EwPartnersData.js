/* =====================================================================
   EDEN BI — EDEN WALK partners & cash flow (admins only)
   Replaces the "EDEN WALK PORTAL" sheet (tab Eden_Walk_New).
   One row per sold unit in "EW Units" (main spreadsheet): price build-up,
   down payment, commission, and installments / guaranteed returns per year.
   Everything else is calculated:
     contract  = down payment + all installments
     net       = contract + returns (returns are negative)
     EDEN      = 64% × (net − commission)     AFROA = 36% × (net − commission)
   Payments received are logged in "EW Collections".
   ===================================================================== */
const EW_UNITS_SHEET_ = 'EW Units';
const EW_COLL_SHEET_ = 'EW Collections';
const EW_PORTAL_ID_ = '1Lr1SqWzD0utKc9EldwSXXyUGVpuMm4fzQ1hG5Ly2xKY';
const EW_PORTAL_TAB_ = 'Eden_Walk_New';
const EW_SHARE_EDEN_ = 0.64;
const EW_COMM_RATE_ = 0.10;
const EW_YEARS_ = [2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035, 2036];
const EW_BASE_HEAD_ = ['Unit Code', 'Category', 'Offer', 'Sale Date', 'Indoor Area', 'Outdoor Area',
  'Basic Indoor Price/m', 'Basic Outdoor Price/m', 'Discount', 'Net Total Price', 'DP %', 'DP Amount', 'DP Paid',
  'Commission', 'Notes', 'Updated By', 'Updated At'];
const EW_COLL_HEAD_ = ['Payment ID', 'Unit Code', 'Date', 'Type', 'Year', 'Amount', 'Method', 'Reference', 'Notes', 'Recorded By', 'Recorded At'];
const EW_CATS_ = ['Semi Cash', 'Return', 'Normal', 'KSA Old'];

function ewHead_(years) {
  const h = EW_BASE_HEAD_.slice();
  (years || EW_YEARS_).forEach(y => { h.push('Installments ' + y); h.push('Return ' + y); });
  return h;
}
function ewAdmin_(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('EDEN WALK partners page is for admins only.');
  return user;
}
/* Who can open the section: admins (edit) + usernames in Script Property
   EW_VIEWERS (view only, default "cfo"). Egypt viewers never, whatever the list says. */
function ewViewers_() {
  const v = PropertiesService.getScriptProperties().getProperty('EW_VIEWERS');
  return (v === null ? 'cfo' : v).split(',').map(x => x.trim().toLowerCase()).filter(Boolean);
}
function ewAccessOf_(user) {
  const role = String(user && user.role || '').trim().toLowerCase();
  if (/egypt/.test(role)) return { see: false, edit: false };
  if (role === 'admin' || role === 'super admin') return { see: true, edit: true };
  const un = String(user && user.username || '').trim().toLowerCase();
  return { see: !!un && ewViewers_().indexOf(un) >= 0, edit: false };
}
function ewViewer_(authToken) {
  const user = validateAuthToken_(authToken);
  if (!ewAccessOf_(user).see) throw new Error('EDEN WALK partners page is not available for this account.');
  return user;
}
function getEwAccess(authToken) {
  const user = validateAuthToken_(authToken);
  return ewAccessOf_(user);
}
function ewWho_(user) { return String(user && (user.name || user.username || user.email) || 'Admin'); }
function ewNum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let t = String(v).replace(/[,\s%]/g, '');
  let neg = false;
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  if (t.charAt(0) === '-') { neg = !neg; t = t.slice(1); }
  if (t === '' || t === '-') return 0;
  const n = Number(t);
  return isFinite(n) ? (neg ? -n : n) : 0;
}
function ewPct_(v) { if (v === '' || v === null || v === undefined) return ''; const n = ewNum_(v); return Math.abs(n) > 1 ? n / 100 : n; }
function ewKey_(u) { return String(u || '').trim().toUpperCase().replace(/\s+/g, ''); }
function ewSheet_(create) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(EW_UNITS_SHEET_);
  if (!sh && create) {
    sh = ss.insertSheet(EW_UNITS_SHEET_);
    const h = ewHead_();
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1); sh.setFrozenColumns(1);
  }
  return sh;
}
function ewYearsOf_(head) {
  return head.map(String).filter(h => /^Installments \d{4}$/.test(h)).map(h => Number(h.slice(-4)));
}
function ewRead_() {
  const sh = ewSheet_(false);
  if (!sh || sh.getLastRow() < 1) return { units: [], years: EW_YEARS_.slice(), head: ewHead_() };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const years = ewYearsOf_(head);
  const ix = {}; head.forEach((h, i) => { ix[h] = i; });
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
  const units = rows.map((r, i) => {
    const g = k => r[ix[k]];
    const plan = {};
    years.forEach(y => { plan[y] = [ewNum_(g('Installments ' + y)), ewNum_(g('Return ' + y))]; });
    return {
      row: i + 2, unit: String(g('Unit Code') || '').trim(), category: String(g('Category') || '').trim(), offer: String(g('Offer') || '').trim(),
      date: instDateOut_(g('Sale Date')), inArea: ewNum_(g('Indoor Area')), outArea: ewNum_(g('Outdoor Area')),
      inPrice: ewNum_(g('Basic Indoor Price/m')), outPrice: ewNum_(g('Basic Outdoor Price/m')), discount: ewPct_(g('Discount')),
      netTotal: ewNum_(g('Net Total Price')), dpPct: ewPct_(g('DP %')), dp: ewNum_(g('DP Amount')), dpPaid: ewNum_(g('DP Paid')),
      commission: ewNum_(g('Commission')), notes: String(g('Notes') || ''), by: String(g('Updated By') || ''), at: instDateOut_(g('Updated At')), plan: plan
    };
  }).filter(u => u.unit);
  return { units: units, years: years, head: head };
}

function ewSettings_() {
  const p = PropertiesService.getScriptProperties();
  const sh = Number(p.getProperty('EW_SHARE_EDEN')), cr = Number(p.getProperty('EW_COMM_RATE'));
  return { share: sh > 0 && sh < 1 ? sh : EW_SHARE_EDEN_, commRate: cr >= 0 && cr < 1 && p.getProperty('EW_COMM_RATE') !== null ? cr : EW_COMM_RATE_ };
}
function saveEwSettings(authToken, req) {
  const user = ewAdmin_(authToken);
  if (req && req.viewers !== undefined) {
    const list = String(req.viewers || '').split(/[,\s]+/).map(x => x.trim().toLowerCase()).filter(Boolean);
    PropertiesService.getScriptProperties().setProperty('EW_VIEWERS', list.join(','));
  }
  const share = ewNum_(req && req.share) / 100, rate = ewNum_(req && req.commRate) / 100;
  if (!(share > 0 && share < 1)) throw new Error('EDEN share must be between 0 and 100 %.');
  if (!(rate >= 0 && rate < 1)) throw new Error('Commission must be between 0 and 100 %.');
  const old = ewSettings_();
  PropertiesService.getScriptProperties().setProperties({ EW_SHARE_EDEN: String(share), EW_COMM_RATE: String(rate) });
  auditLog_(user, 'EW settings changed', { project: 'EDEN WALK', details: 'EDEN ' + Math.round(old.share * 1000) / 10 + '% → ' + Math.round(share * 1000) / 10 + '%, commission ' + Math.round(old.commRate * 1000) / 10 + '% → ' + Math.round(rate * 1000) / 10 + '%' });
  return getEwPartners(authToken);
}

/* Everything the page needs. */
function getEwPartners(authToken) {
  const user = ewViewer_(authToken), acc = ewAccessOf_(user);
  const d = ewRead_();
  const coll = instRows_(instSheet_(EW_COLL_SHEET_, EW_COLL_HEAD_, false)).filter(r => r[0]).map(r => ({
    id: String(r[0]), unit: String(r[1] || ''), date: instDateOut_(r[2]), type: String(r[3] || ''), year: Number(r[4]) || '',
    amount: ewNum_(r[5]), method: String(r[6] || ''), ref: String(r[7] || ''), notes: String(r[8] || ''), by: String(r[9] || ''), at: instDateOut_(r[10])
  }));
  const st = ewSettings_();
  return { canEdit: acc.edit, viewers: acc.edit ? ewViewers_() : [], units: d.units, years: d.years, collections: coll, share: st.share, commRate: st.commRate, categories: EW_CATS_, imported: d.units.length > 0 };
}

/* One-time copy from the old portal sheet (tab Eden_Walk_New). Refuses if units already exist unless replace=true. */
function importEwFromPortal(authToken, req) {
  const user = ewAdmin_(authToken);
  req = req || {};
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const cur = ewRead_();
    if (cur.units.length && !req.replace) throw new Error('EDEN WALK units are already in the system (' + cur.units.length + ').');
    const src = SpreadsheetApp.openById(EW_PORTAL_ID_).getSheetByName(EW_PORTAL_TAB_);
    if (!src) throw new Error('Tab "' + EW_PORTAL_TAB_ + '" not found in the EDEN WALK PORTAL sheet.');
    const vals = src.getDataRange().getValues();
    const top = vals[0].map(String), sub = (vals[1] || []).map(String);
    // year columns: a 4-digit year in the top header, "Installments" under it, "Return" next
    const yearCols = [];
    top.forEach((h, i) => { const y = Number(String(h).trim()); if (y >= 2000 && y < 2100 && /install/i.test(sub[i] || '')) yearCols.push({ y: y, c: i }); });
    const years = yearCols.map(x => x.y);
    const sumStart = top.findIndex(h => /summary/i.test(h));
    if (sumStart < 0 || !years.length) throw new Error('The portal sheet layout is not the expected one (Summary / year columns not found).');
    const out = [], catOf = t => /semi/i.test(t) ? 'Semi Cash' : /return/i.test(t) ? 'Return' : /ksa/i.test(t) ? 'KSA Old' : /normal/i.test(t) ? 'Normal' : '';
    let pending = [];
    for (let i = 2; i < vals.length; i++) {
      const r = vals[i], code = String(r[0] || '').trim(), label = String(r[8] || '');
      if (/^grand total/i.test(code)) break;
      if (!code) {
        if (/total\s*-/i.test(label)) { const c = catOf(label); pending.forEach(x => { x[1] = c; }); pending = []; }
        continue;
      }
      const line = [code, '', label.trim(), r[16] instanceof Date ? r[16] : (dealDate_(r[16]) || r[16]), ewNum_(r[1]), ewNum_(r[2]), ewNum_(r[3]), ewNum_(r[4]),
        ewPct_(r[9]), ewNum_(r[14]), ewPct_(r[17]), ewNum_(r[18]), ewNum_(r[19]), ewNum_(r[sumStart + 4]), '', ewWho_(user) + ' (import)', new Date()];
      yearCols.forEach(x => { line.push(ewNum_(r[x.c])); line.push(ewNum_(r[x.c + 1])); });
      out.push(line); pending.push(line);
    }
    if (!out.length) throw new Error('No units found in the portal sheet.');
    let sh = ewSheet_(false);
    const head = ewHead_(years);
    if (!sh) { sh = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(EW_UNITS_SHEET_); }
    sh.clear();
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.getRange(2, 1, out.length, head.length).setValues(out);
    sh.setFrozenRows(1); sh.setFrozenColumns(1);
    sh.getRange(2, 4, out.length, 1).setNumberFormat('yyyy-mm-dd');
    auditLog_(user, 'EW units imported', { project: 'EDEN WALK', details: out.length + ' units from EDEN WALK PORTAL / ' + EW_PORTAL_TAB_ });
  } finally { lock.releaseLock(); }
  return getEwPartners(authToken);
}

/* Add (row empty) or update one unit. */
function saveEwUnit(authToken, req) {
  const user = ewAdmin_(authToken);
  req = req || {};
  const code = String(req.unit || '').trim();
  if (!code) throw new Error('Unit code is missing.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let isNew = false;
  try {
    const sh = ewSheet_(true);
    const d = ewRead_();
    const head = d.head, years = d.years.slice();
    // new years in the plan get their own columns
    Object.keys(req.plan || {}).map(Number).filter(y => y >= 2000 && y < 2100 && years.indexOf(y) < 0).sort().forEach(y => {
      sh.getRange(1, head.length + 1, 1, 2).setValues([['Installments ' + y, 'Return ' + y]]).setFontWeight('bold');
      head.push('Installments ' + y); head.push('Return ' + y); years.push(y);
    });
    const ix = {}; head.forEach((h, i) => { ix[h] = i; });
    const same = d.units.find(u => ewKey_(u.unit) === ewKey_(code));
    const target = req.row ? d.units.find(u => u.row === Number(req.row)) : null;
    if (same && (!target || same.row !== target.row)) throw new Error('Unit ' + code + ' is already in the list.');
    if (req.row && !target) throw new Error('This unit was changed by someone else — reload the page.');
    isNew = !target;
    const line = target ? sh.getRange(target.row, 1, 1, head.length).getValues()[0] : new Array(head.length).fill('');
    while (line.length < head.length) line.push('');
    const set = (k, v) => { if (ix[k] !== undefined) line[ix[k]] = v; };
    set('Unit Code', code); set('Category', String(req.category || '')); set('Offer', String(req.offer || ''));
    set('Sale Date', instDateIn_(req.date)); set('Indoor Area', ewNum_(req.inArea)); set('Outdoor Area', ewNum_(req.outArea));
    set('Basic Indoor Price/m', ewNum_(req.inPrice)); set('Basic Outdoor Price/m', ewNum_(req.outPrice));
    set('Discount', req.discount === '' || req.discount == null ? '' : ewNum_(req.discount) / 100);
    set('Net Total Price', ewNum_(req.netTotal)); set('DP %', req.dpPct === '' || req.dpPct == null ? '' : ewNum_(req.dpPct) / 100);
    set('DP Amount', ewNum_(req.dp)); set('DP Paid', ewNum_(req.dpPaid)); set('Commission', ewNum_(req.commission));
    set('Notes', String(req.notes || '')); set('Updated By', ewWho_(user)); set('Updated At', new Date());
    years.forEach(y => { const p = (req.plan || {})[y] || [0, 0]; set('Installments ' + y, ewNum_(p[0])); set('Return ' + y, -Math.abs(ewNum_(p[1]))); });
    if (target) sh.getRange(target.row, 1, 1, head.length).setValues([line]);
    else sh.appendRow(line);
    auditLog_(user, isNew ? 'EW sale added' : 'EW unit updated', { project: 'EDEN WALK', unit: code, details: (req.offer || '') + ' · DP ' + ewNum_(req.dp) + ' · paid ' + ewNum_(req.dpPaid) });
  } finally { lock.releaseLock(); }
  return getEwPartners(authToken);
}

function deleteEwUnit(authToken, req) {
  const user = ewAdmin_(authToken);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = ewRead_();
    const t = d.units.find(u => u.row === Number(req && req.row) && ewKey_(u.unit) === ewKey_(req && req.unit));
    if (!t) throw new Error('Unit not found — reload the page.');
    ewSheet_(false).deleteRow(t.row);
    auditLog_(user, 'EW unit deleted', { project: 'EDEN WALK', unit: t.unit, details: t.offer });
  } finally { lock.releaseLock(); }
  return getEwPartners(authToken);
}

/* Money received. A down-payment receipt also raises "DP Paid" on the unit. */
function recordEwCollection(authToken, req) {
  const user = ewAdmin_(authToken);
  req = req || {};
  const amount = ewNum_(req.amount);
  if (!(amount > 0)) throw new Error('Enter the amount.');
  const type = /install/i.test(req.type) ? 'Installment' : 'Down payment';
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = ewRead_();
    const u = d.units.find(x => ewKey_(x.unit) === ewKey_(req.unit));
    if (!u) throw new Error('Unit not found.');
    const sh = instSheet_(EW_COLL_SHEET_, EW_COLL_HEAD_, true);
    const id = 'EWP-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const date = instDateIn_(req.date) || new Date();
    sh.appendRow([id, u.unit, date, type, type === 'Installment' ? (Number(req.year) || date.getFullYear()) : '', amount,
      String(req.method || ''), String(req.ref || ''), String(req.notes || ''), ewWho_(user), new Date()]);
    if (type === 'Down payment') {
      const head = d.head, c = head.indexOf('DP Paid') + 1;
      if (c > 0) ewSheet_(false).getRange(u.row, c).setValue(u.dpPaid + amount);
    }
    auditLog_(user, 'EW payment recorded', { project: 'EDEN WALK', unit: u.unit, details: type + ' ' + amount + (req.ref ? ' · ' + req.ref : '') });
  } finally { lock.releaseLock(); }
  return getEwPartners(authToken);
}

function deleteEwCollection(authToken, req) {
  const user = ewAdmin_(authToken);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = instSheet_(EW_COLL_SHEET_, EW_COLL_HEAD_, false);
    const rows = instRows_(sh);
    const i = rows.findIndex(r => String(r[0]) === String(req && req.id));
    if (i < 0) throw new Error('Payment not found.');
    const r = rows[i];
    if (String(r[3]) === 'Down payment') {
      const d = ewRead_(), u = d.units.find(x => ewKey_(x.unit) === ewKey_(r[1])), c = d.head.indexOf('DP Paid') + 1;
      if (u && c > 0) ewSheet_(false).getRange(u.row, c).setValue(Math.max(0, u.dpPaid - ewNum_(r[5])));
    }
    sh.deleteRow(i + 2);
    auditLog_(user, 'EW payment deleted', { project: 'EDEN WALK', unit: String(r[1]), details: String(r[3]) + ' ' + ewNum_(r[5]) });
  } finally { lock.releaseLock(); }
  return getEwPartners(authToken);
}
