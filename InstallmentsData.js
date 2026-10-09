/************************************************
 * EDEN BI V4 — Installments.js
 *
 * Payment schedule per unit + the payments clients make.
 *  - "Installments" tab: one row per installment (Down Payment, Q1…Qn,
 *    Maintenance) of a unit.
 *  - "Installment Payments" tab: every payment recorded, one row each.
 *    What is paid on an installment = the sum of its payments, so a
 *    payment can be partial and nothing is ever overwritten.
 * Both tabs live in the main spreadsheet and are created on first use.
 * Reading is open to signed-in users; changes are for admins only.
 ************************************************/

const INST_SHEET_ = 'Installments';
const INST_PAY_SHEET_ = 'Installment Payments';
const INST_HEAD_ = ['Key', 'Project', 'Unit Code', 'Client', 'Deal Row', 'No', 'Name', 'Due Date', 'Amount', 'Currency', 'Created By', 'Created At'];
const INST_PAY_HEAD_ = ['Payment ID', 'Key', 'Project', 'Unit Code', 'No', 'Amount', 'Paid Date', 'Method', 'Receipt', 'Notes', 'Recorded By', 'Recorded At'];

function instIsAdmin_(user) {
  const role = String(user && user.role || '').trim().toLowerCase();
  return role === 'admin' || role === 'super admin';
}
function instKey_(project, unitCode) {
  return dealProjectKey_(project) + '|' + String(unitCode || '').trim().replace(/\.+$/, '').toUpperCase();
}
function instSheet_(name, head, create) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh && create) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function instTz_() { return Session.getScriptTimeZone() || 'Africa/Cairo'; }
function instDateOut_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, instTz_(), 'yyyy-MM-dd');
  return String(v || '').trim();
}
function instDateIn_(v) {
  const t = String(v || '').trim();
  const m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  const d = dealDate_(t);
  return d || '';
}
function instRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

/* Schedule + payments of one unit. */
function getInstallmentSchedule(authToken, project, unitCode, dealRow) {
  const user = validateAuthToken_(authToken);
  const key = instKey_(project, unitCode);
  const sched = instRows_(instSheet_(INST_SHEET_, INST_HEAD_, false))
    .filter(r => String(r[0]) === key)
    .map(r => ({ no: String(r[5]), name: String(r[6]), due: instDateOut_(r[7]), amount: Number(r[8]) || 0, currency: String(r[9] || ''), client: String(r[3] || ''), dealRow: Number(r[4]) || 0 }));
  const pays = instRows_(instSheet_(INST_PAY_SHEET_, INST_PAY_HEAD_, false))
    .filter(r => String(r[1]) === key)
    .map(r => ({ id: String(r[0]), no: String(r[4]), amount: Number(r[5]) || 0, date: instDateOut_(r[6]), method: String(r[7] || ''), receipt: String(r[8] || ''), notes: String(r[9] || ''), by: String(r[10] || ''), at: instDateOut_(r[11]) }));
  return { key: key, schedule: sched, payments: pays, canEdit: instIsAdmin_(user), plan: instDealPlan_(dealRow, unitCode) };
}

/* The payment plan as written in the deal's row of the Deals sheet (read live, not from cache). */
function instDealPlan_(dealRow, unitCode) {
  const row = Number(dealRow);
  if (!row || row < DATA_START_ROW) return null;
  try {
    const sheet = getMainDealsSheet_();
    if (row > sheet.getLastRow()) return null;
    const cols = mainDealsColumns_(sheet);
    const vals = sheet.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
    const at = names => { const c = mainDealsCol_(cols, names); return c ? String(vals[c - 1] || '').trim() : ''; };
    if (unitCode && at(['Unit Code']).toUpperCase().replace(/\.+$/, '') !== String(unitCode).trim().toUpperCase().replace(/\.+$/, '')) return null;
    return {
      finalPrice: at(['Final Price']),
      dpPercent: at(['DP %', 'DP%']),
      dpDate: at(['down payment|date']),
      contractDate: at(['contract details|date']),
      maintenancePercent: at(['Maintenance %']),
      maintenanceAmount: at(['Maintenance Amount']),
      planType: at(['payment plan|type of payment plan', 'Type Of Payment Plan']),
      planPeriod: at(['payment plan|installments period', 'Installments Period']),
      planCode: at(['payment plan|installment plan', 'Installment Plan']),
      planEvery: at(['payment plan|period type', 'Period Type'])
    };
  } catch (e) { return null; }
}

/* Admin: create or replace the schedule of a unit (payments are kept). */
function saveInstallmentSchedule(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!instIsAdmin_(user)) throw new Error('Only admins can change payment schedules.');
  req = req || {};
  const rows = Array.isArray(req.rows) ? req.rows : [];
  if (!req.project || !req.unitCode) throw new Error('Unit is missing.');
  if (!rows.length) throw new Error('The schedule is empty.');
  if (rows.length > 400) throw new Error('Too many installments.');
  const key = instKey_(req.project, req.unitCode);
  const now = new Date(), by = user.name || user.username;
  const out = rows.map((r, i) => {
    const due = instDateIn_(r.due);
    const amt = Number(r.amount);
    if (!due) throw new Error('Row ' + (i + 1) + ': due date is missing.');
    if (!(amt >= 0)) throw new Error('Row ' + (i + 1) + ': amount is not valid.');
    return [key, dealText_(req.project), dealText_(req.unitCode), dealText_(req.client), Number(req.dealRow) || '',
      String(r.no || i), dealText_(r.name) || ('#' + (i + 1)), due, Math.round(amt * 100) / 100, dealText_(req.currency), by, now];
  });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { throw new Error('Busy, try again.'); }
  try {
    const sh = instSheet_(INST_SHEET_, INST_HEAD_, true);
    // Remove the old rows of this unit (bottom-up), then append the new ones.
    const last = sh.getLastRow();
    if (last >= 2) {
      const keys = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = keys.length - 1; i >= 0; i--) if (String(keys[i][0]) === key) sh.deleteRow(i + 2);
    }
    sh.getRange(sh.getLastRow() + 1, 1, out.length, INST_HEAD_.length).setValues(out);
    sh.getRange(2, 8, Math.max(1, sh.getLastRow() - 1), 1).setNumberFormat('dd-MMM-yyyy');
    sh.getRange(2, 9, Math.max(1, sh.getLastRow() - 1), 1).setNumberFormat('#,##0');

    // Optional: carry the down payment already recorded in Deals as a first payment.
    const carry = Number(req.carryPaid) || 0;
    if (carry > 0) {
      const pays = instRows_(instSheet_(INST_PAY_SHEET_, INST_PAY_HEAD_, false)).filter(r => String(r[1]) === key);
      if (!pays.length) {
        instAppendPayment_(key, req.project, req.unitCode, String(req.carryNo || '0'), carry, instDateIn_(req.carryDate) || now,
          'Deals sheet', '', 'Down payment paid before the schedule was created', by);
      }
    }
  } finally {
    lock.releaseLock();
  }
  const main = rows.filter(r => !/^(down payment|maintenance)$/i.test(String(r.name || ''))).length;
  const total = rows.reduce((t, r) => t + (Number(r.amount) || 0), 0);
  auditLog_(user, 'Payment schedule saved', { project: req.project, unit: req.unitCode, client: req.client,
    details: main + ' installments · total ' + Math.round(total).toLocaleString('en-US') + ' ' + dealText_(req.currency) + (Number(req.carryPaid) > 0 ? ' · down payment carried ' + Number(req.carryPaid).toLocaleString('en-US') : '') });
  return getInstallmentSchedule(authToken, req.project, req.unitCode);
}

function instAppendPayment_(key, project, unitCode, no, amount, date, method, receipt, notes, by) {
  const sh = instSheet_(INST_PAY_SHEET_, INST_PAY_HEAD_, true);
  const id = 'P' + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5).toUpperCase();
  sh.appendRow([id, key, dealText_(project), dealText_(unitCode), String(no), Math.round(Number(amount) * 100) / 100, date,
    dealText_(method), dealText_(receipt), dealText_(notes), by, new Date()]);
  const r = sh.getLastRow();
  sh.getRange(r, 7).setNumberFormat('dd-MMM-yyyy');
  sh.getRange(r, 6).setNumberFormat('#,##0');
  return id;
}

/* Admin: record a payment against one installment. */
function recordInstallmentPayment(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!instIsAdmin_(user)) throw new Error('Only admins can record payments.');
  req = req || {};
  const amount = Number(String(req.amount || '').replace(/,/g, ''));
  if (!(amount > 0)) throw new Error('Enter the amount paid.');
  const date = instDateIn_(req.date);
  if (!date) throw new Error('Enter the payment date.');
  const key = instKey_(req.project, req.unitCode);
  const has = instRows_(instSheet_(INST_SHEET_, INST_HEAD_, false)).some(r => String(r[0]) === key && String(r[5]) === String(req.no));
  if (!has) throw new Error('This installment was not found. Refresh and try again.');
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { throw new Error('Busy, try again.'); }
  try {
    instAppendPayment_(key, req.project, req.unitCode, req.no, amount, date, req.method, req.receipt, req.notes, user.name || user.username);
  } finally {
    lock.releaseLock();
  }
  const inst = instRows_(instSheet_(INST_SHEET_, INST_HEAD_, false)).find(r => String(r[0]) === key && String(r[5]) === String(req.no));
  auditLog_(user, 'Payment recorded', { project: req.project, unit: req.unitCode, client: inst ? String(inst[3] || '') : '',
    details: Math.round(amount).toLocaleString('en-US') + ' on ' + (inst ? String(inst[6]) : '#' + req.no) + ' · ' + Utilities.formatDate(date, instTz_(), 'dd-MMM-yyyy') +
      (req.method ? ' · ' + dealText_(req.method) : '') + (req.receipt ? ' · Receipt ' + dealText_(req.receipt) : ''), ref: '' });
  return getInstallmentSchedule(authToken, req.project, req.unitCode);
}

/* Admin: delete a payment entered by mistake. */
function deleteInstallmentPayment(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!instIsAdmin_(user)) throw new Error('Only admins can delete payments.');
  req = req || {};
  const sh = instSheet_(INST_PAY_SHEET_, INST_PAY_HEAD_, false);
  if (!sh) throw new Error('Payment not found.');
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { throw new Error('Busy, try again.'); }
  try {
    const ids = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
    const i = ids.findIndex(r => String(r[0]) === String(req.id));
    if (i < 0) throw new Error('Payment not found.');
    const old = sh.getRange(i + 2, 1, 1, INST_PAY_HEAD_.length).getValues()[0];
    sh.deleteRow(i + 2);
    auditLog_(user, 'Payment deleted', { project: req.project, unit: req.unitCode,
      details: Math.round(Number(old[5]) || 0).toLocaleString('en-US') + ' (installment ' + old[4] + ', paid ' + instDateOut_(old[6]) + (old[7] ? ', ' + old[7] : '') + (old[8] ? ', receipt ' + old[8] : '') + ')', ref: String(old[0]) });
  } finally {
    lock.releaseLock();
  }
  return getInstallmentSchedule(authToken, req.project, req.unitCode);
}

/* Every installment with what is paid on it (for the Collections page). */
function getInstallmentsOverview(authToken) {
  validateAuthToken_(authToken);
  const paid = {};
  instRows_(instSheet_(INST_PAY_SHEET_, INST_PAY_HEAD_, false)).forEach(r => {
    const k = String(r[1]) + '#' + String(r[4]);
    paid[k] = (paid[k] || 0) + (Number(r[5]) || 0);
  });
  const rows = instRows_(instSheet_(INST_SHEET_, INST_HEAD_, false)).map(r => ({
    key: String(r[0]), project: String(r[1]), unitCode: String(r[2]), client: String(r[3]), no: String(r[5]), name: String(r[6]),
    due: instDateOut_(r[7]), amount: Number(r[8]) || 0, currency: String(r[9] || ''), paid: paid[String(r[0]) + '#' + String(r[5])] || 0
  }));
  return { rows: rows, generatedAt: new Date().toISOString() };
}
