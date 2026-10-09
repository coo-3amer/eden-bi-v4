/************************************************
 * EDEN BI V4 — BrokerCommData.js
 *
 * Commissions owed to broker companies on the deals they brought.
 *  - "Broker Commission Rates": one row per company, the rate % typed by
 *    the admin (rates differ per company, nothing is assumed).
 *  - "Broker Deal Commission": optional per-deal override (rate % or a
 *    fixed amount) when one deal was agreed differently.
 *  - "Broker Commission Payouts": every payment made to a company, one row
 *    each, linked to the deal it pays for. Nothing is overwritten.
 * The deals themselves (company, price, status) come from the Deals sheet
 * already loaded in the dashboard. Admins only — these are company money.
 ************************************************/

const BC_RATES_SHEET_ = 'Broker Commission Rates';
const BC_DEAL_SHEET_ = 'Broker Deal Commission';
const BC_PAY_SHEET_ = 'Broker Commission Payouts';
const BC_RATES_HEAD_ = ['Company', 'Rate %', 'Notes', 'Updated By', 'Updated At'];
const BC_DEAL_HEAD_ = ['Deal Key', 'Deal Row', 'Project', 'Unit Code', 'Company', 'Rate %', 'Fixed Amount', 'Notes', 'Updated By', 'Updated At'];
const BC_PAY_HEAD_ = ['Payout ID', 'Company', 'Deal Key', 'Project', 'Unit Code', 'Client', 'Amount', 'Currency', 'Paid Date', 'Method', 'Reference', 'Notes', 'Recorded By', 'Recorded At'];

function bcAdmin_(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Broker commissions are for admins only.');
  return user;
}
function bcWho_(user) { return String(user && (user.name || user.username || user.email) || 'Admin'); }
function bcCompanyKey_(name) { return String(name || '').trim().replace(/\s+/g, ' ').toLowerCase(); }
function bcNum_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(String(v).replace(/[,%\s]/g, ''));
  return isFinite(n) ? n : null;
}

/* Everything the commissions page needs (rates, per-deal overrides, payouts). */
function getBrokerCommissions(authToken) {
  const user = bcAdmin_(authToken);
  const rates = {};
  instRows_(instSheet_(BC_RATES_SHEET_, BC_RATES_HEAD_, false)).forEach(r => {
    const k = bcCompanyKey_(r[0]);
    if (k) rates[k] = { company: String(r[0]).trim(), rate: bcNum_(r[1]), notes: String(r[2] || ''), by: String(r[3] || ''), at: instDateOut_(r[4]) };
  });
  const deals = {};
  instRows_(instSheet_(BC_DEAL_SHEET_, BC_DEAL_HEAD_, false)).forEach(r => {
    const k = String(r[0] || '').trim();
    if (k) deals[k] = { dealRow: Number(r[1]) || 0, project: String(r[2] || ''), unitCode: String(r[3] || ''), company: String(r[4] || ''), rate: bcNum_(r[5]), amount: bcNum_(r[6]), notes: String(r[7] || ''), by: String(r[8] || ''), at: instDateOut_(r[9]) };
  });
  const payouts = instRows_(instSheet_(BC_PAY_SHEET_, BC_PAY_HEAD_, false)).filter(r => r[0]).map(r => ({
    id: String(r[0]), company: String(r[1] || ''), dealKey: String(r[2] || ''), project: String(r[3] || ''), unitCode: String(r[4] || ''),
    client: String(r[5] || ''), amount: Number(r[6]) || 0, currency: String(r[7] || ''), date: instDateOut_(r[8]),
    method: String(r[9] || ''), ref: String(r[10] || ''), notes: String(r[11] || ''), by: String(r[12] || ''), at: instDateOut_(r[13])
  }));
  return { rates: rates, deals: deals, payouts: payouts, by: bcWho_(user) };
}

/* Sets (or clears, with an empty rate) the commission rate of one company. */
function saveBrokerCommissionRate(authToken, req) {
  const user = bcAdmin_(authToken);
  req = req || {};
  const company = String(req.company || '').trim();
  if (!company) throw new Error('Company is missing.');
  const rate = bcNum_(req.rate);
  if (rate !== null && (rate < 0 || rate > 100)) throw new Error('The rate must be between 0 and 100 %.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = instSheet_(BC_RATES_SHEET_, BC_RATES_HEAD_, true);
    const rows = instRows_(sh);
    const i = rows.findIndex(r => bcCompanyKey_(r[0]) === bcCompanyKey_(company));
    const old = i >= 0 ? bcNum_(rows[i][1]) : null;
    const line = [company, rate === null ? '' : rate, String(req.notes || '').trim(), bcWho_(user), new Date()];
    if (i >= 0) sh.getRange(i + 2, 1, 1, line.length).setValues([line]);
    else sh.appendRow(line);
    auditLog_(user, 'Broker commission rate', { details: company + ': ' + (old === null ? '—' : old + '%') + ' → ' + (rate === null ? '—' : rate + '%') });
  } finally { lock.releaseLock(); }
  return getBrokerCommissions(authToken);
}

/* Per-deal override: a different rate % or a fixed amount for one deal (empty both = back to the company rate). */
function saveBrokerDealCommission(authToken, req) {
  const user = bcAdmin_(authToken);
  req = req || {};
  const key = instKey_(req.project, req.unitCode);
  if (!req.project || !req.unitCode) throw new Error('Deal is missing.');
  const rate = bcNum_(req.rate), amount = bcNum_(req.amount);
  if (rate !== null && (rate < 0 || rate > 100)) throw new Error('The rate must be between 0 and 100 %.');
  if (amount !== null && amount < 0) throw new Error('The amount cannot be negative.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = instSheet_(BC_DEAL_SHEET_, BC_DEAL_HEAD_, true);
    const rows = instRows_(sh);
    const i = rows.findIndex(r => String(r[0]) === key);
    if (rate === null && amount === null && !String(req.notes || '').trim()) {
      if (i >= 0) sh.deleteRow(i + 2);
    } else {
      const line = [key, Number(req.dealRow) || '', req.project, req.unitCode, String(req.company || ''), rate === null ? '' : rate, amount === null ? '' : amount, String(req.notes || '').trim(), bcWho_(user), new Date()];
      if (i >= 0) sh.getRange(i + 2, 1, 1, line.length).setValues([line]);
      else sh.appendRow(line);
    }
    auditLog_(user, 'Broker commission (deal)', { project: req.project, unit: req.unitCode,
      details: String(req.company || '') + ': ' + (rate === null && amount === null ? 'back to the company rate' : (amount !== null ? 'fixed ' + amount : rate + '%')) });
  } finally { lock.releaseLock(); }
  return getBrokerCommissions(authToken);
}

/* Records a payment made to a broker company for one deal. */
function recordBrokerPayout(authToken, req) {
  const user = bcAdmin_(authToken);
  req = req || {};
  const amount = bcNum_(req.amount);
  if (!amount || amount <= 0) throw new Error('Enter the amount paid.');
  if (!req.company) throw new Error('Company is missing.');
  if (!req.project || !req.unitCode) throw new Error('Deal is missing.');
  const date = instDateIn_(req.date) || new Date();
  const id = 'BP-' + Utilities.formatDate(new Date(), instTz_(), 'yyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    instSheet_(BC_PAY_SHEET_, BC_PAY_HEAD_, true).appendRow([id, String(req.company).trim(), instKey_(req.project, req.unitCode), req.project, req.unitCode,
      String(req.client || ''), amount, String(req.currency || ''), date, String(req.method || ''), String(req.ref || ''), String(req.notes || ''), bcWho_(user), new Date()]);
  } finally { lock.releaseLock(); }
  auditLog_(user, 'Broker commission paid', { project: req.project, unit: req.unitCode, client: req.client,
    details: String(req.company) + ': ' + amount + ' ' + String(req.currency || '') + (req.method ? ' · ' + req.method : '') + (req.ref ? ' · ' + req.ref : ''), ref: id });
  return getBrokerCommissions(authToken);
}

function deleteBrokerPayout(authToken, req) {
  const user = bcAdmin_(authToken);
  const id = String(req && req.id || '').trim();
  if (!id) throw new Error('Payout is missing.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = instSheet_(BC_PAY_SHEET_, BC_PAY_HEAD_, false);
    const rows = instRows_(sh);
    const i = rows.findIndex(r => String(r[0]) === id);
    if (i < 0) throw new Error('This payout was not found (maybe already deleted).');
    const r = rows[i];
    sh.deleteRow(i + 2);
    auditLog_(user, 'Broker commission payout deleted', { project: r[3], unit: r[4], client: r[5], details: r[1] + ': ' + r[6] + ' ' + r[7], ref: id });
  } finally { lock.releaseLock(); }
  return getBrokerCommissions(authToken);
}
