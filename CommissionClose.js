/************************************************
 * EDEN BI V4 — CommissionClose.js
 *
 * Monthly close of internal commissions. Closing a month saves a snapshot
 * of every person's commission (and the deals behind it) in the hidden
 * sheet "_Comm Closings". A closed month shows the snapshot — later edits
 * to deals, rates or manual choices no longer change it. Each person can
 * then be marked as paid. Admins only. Reopening deletes the snapshot.
 ************************************************/

const COMM_CLOSE_SHEET_ = '_Comm Closings';
const COMM_CLOSE_HEAD_ = ['Month', 'Name', 'Dept', 'Role', 'Target', 'Credited', 'Achievement', 'Commission', 'Deals', 'Sales', 'How', 'Details', 'Paid', 'Paid At', 'Paid By', 'Closed By', 'Closed At', 'Label'];

function commCloseSheet_() {
  const ss = targetsSpreadsheet_();
  let sh = ss.getSheetByName(COMM_CLOSE_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(COMM_CLOSE_SHEET_);
    sh.getRange(1, 1, 1, COMM_CLOSE_HEAD_.length).setValues([COMM_CLOSE_HEAD_]).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}
function commCloseRows_() {
  const sh = commCloseSheet_(), last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, COMM_CLOSE_HEAD_.length).getValues();
}
function commMonthKey_(m) {
  const s = String(m || '').trim();
  if (!/^\d{4}-\d{2}$/.test(s)) throw new Error('Pick one whole month to close.');
  return s;
}
function commIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

/* Snapshot of a month: {month, closed, closedBy, closedAt, label, people:[…]} */
function getCommClosing(authToken, month) {
  commAdmin_(authToken);
  const m = commMonthKey_(month);
  const rows = commCloseRows_().filter(r => String(r[0]) === m);
  if (!rows.length) return { month: m, closed: false, people: [] };
  return {
    month: m, closed: true, closedBy: String(rows[0][15] || ''), closedAt: commIso_(rows[0][16]), label: String(rows[0][17] || ''),
    people: rows.map(r => {
      let details = [];
      try { details = JSON.parse(String(r[11] || '[]')); } catch (e) {}
      return { name: String(r[1]), dept: String(r[2]), role: String(r[3]), target: Number(r[4]) || 0, credited: Number(r[5]) || 0, ach: Number(r[6]) || 0,
        comm: Number(r[7]) || 0, n: Number(r[8]) || 0, sales: Number(r[9]) || 0, how: String(r[10] || ''), details: details,
        paid: r[12] === true || String(r[12]).toUpperCase() === 'TRUE', paidAt: commIso_(r[13]), paidBy: String(r[14] || '') };
    })
  };
}

/* Months already closed (for the month picker hints). */
function listCommClosings(authToken) {
  commAdmin_(authToken);
  const out = {};
  commCloseRows_().forEach(r => {
    const m = String(r[0]); if (!m) return;
    const o = out[m] || (out[m] = { month: m, people: 0, total: 0, paid: 0 });
    o.people++; o.total += Number(r[7]) || 0; if (r[12] === true || String(r[12]).toUpperCase() === 'TRUE') o.paid++;
  });
  return Object.keys(out).sort().reverse().map(k => out[k]);
}

function closeCommMonth(authToken, req) {
  const user = commAdmin_(authToken);
  req = req || {};
  const m = commMonthKey_(req.month);
  const people = (Array.isArray(req.people) ? req.people : []).filter(p => p && String(p.name || '').trim());
  if (!people.length) throw new Error('Nothing to close for this month.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (commCloseRows_().some(r => String(r[0]) === m)) throw new Error('This month is already closed.');
    const now = new Date(), who = (user && (user.name || user.username)) || '';
    const num = v => { const n = Number(v); return isFinite(n) ? Math.round(n * 100) / 100 : 0; };
    const out = people.map(p => {
      let det = JSON.stringify(Array.isArray(p.details) ? p.details : []);
      if (det.length > 45000) det = det.slice(0, 45000);
      return [m, String(p.name).trim(), String(p.dept || 'Sales'), String(p.role || ''), num(p.target), num(p.credited), num(p.ach), num(p.comm),
        num(p.n), num(p.sales), String(p.how || ''), det, false, '', '', who, now, String(req.label || '').slice(0, 300)];
    });
    const sh = commCloseSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, out.length, COMM_CLOSE_HEAD_.length).setValues(out);
    const total = out.reduce((a, r) => a + (Number(r[7]) || 0), 0);
    auditLog_(user, 'Commission month closed', { details: m + ' · ' + out.length + ' people · ' + Math.round(total).toLocaleString('en-US') + ' EGP' + (req.label ? ' · ' + req.label : '') });
  } finally { lock.releaseLock(); }
  return getCommClosing(authToken, m);
}

function reopenCommMonth(authToken, month) {
  const user = commAdmin_(authToken);
  const m = commMonthKey_(month);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = commCloseSheet_(), rows = commCloseRows_();
    if (rows.some(r => String(r[0]) === m && (r[12] === true || String(r[12]).toUpperCase() === 'TRUE')))
      throw new Error('Some people are already marked as paid in this month. Untick “Paid” first.');
    const keep = rows.filter(r => String(r[0]) !== m), removed = rows.length - keep.length;
    if (keep.length) sh.getRange(2, 1, keep.length, COMM_CLOSE_HEAD_.length).setValues(keep);
    if (removed) sh.getRange(2 + keep.length, 1, removed, COMM_CLOSE_HEAD_.length).clearContent();
    auditLog_(user, 'Commission month reopened', { details: m + ' · ' + removed + ' rows removed' });
  } finally { lock.releaseLock(); }
  return getCommClosing(authToken, m);
}

function setCommPaid(authToken, req) {
  const user = commAdmin_(authToken);
  req = req || {};
  const m = commMonthKey_(req.month), names = [].concat(req.names || req.name || []).map(n => String(n || '').trim().toLowerCase()).filter(Boolean);
  const dept = req.dept ? String(req.dept) : '';
  const paid = !!req.paid;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = commCloseSheet_(), rows = commCloseRows_(), now = new Date(), who = (user && (user.name || user.username)) || '';
    let n = 0;
    rows.forEach((r, i) => {
      if (String(r[0]) !== m || names.indexOf(String(r[1]).trim().toLowerCase()) < 0) return;
      if (dept && String(r[2]) !== dept) return;
      sh.getRange(i + 2, 13, 1, 3).setValues([[paid, paid ? now : '', paid ? who : '']]);
      n++;
    });
    if (!n) throw new Error('Not found in the closed month.');
    auditLog_(user, paid ? 'Commission marked paid' : 'Commission marked unpaid', { details: m + ' · ' + names.join(', ') });
  } finally { lock.releaseLock(); }
  return getCommClosing(authToken, m);
}
