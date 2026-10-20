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
const COMM_DEPTS_ = ['Sales', 'Sales Management', 'Customer Service / Contracts', 'Finance / Collection', 'Marketing', 'HR', 'Secretary', 'Buffet', 'Other'];

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
