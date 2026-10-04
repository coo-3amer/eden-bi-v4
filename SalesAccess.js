/************************************************
 * EDEN BI V4 — SalesAccess.js
 *
 * Role-based data scope for the sales team:
 *   Sales          -> only deals where the user is the salesperson or the
 *                     shared partner, and only their own targets.
 *   Sales Manager  -> deals of their team (Sales Manager column), plus their
 *                     own / shared deals, and the targets of their team.
 * The user's Name in the Users sheet must match the name used in the Deals
 * sheet (Sales Name / Sales Manager). Filtering happens on the server, so
 * nothing outside the user's scope is ever sent to the browser.
 ************************************************/

function salesRoleKey_(user) {
  return String((user && user.role) || '').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
}
function isSalesUser_(user) { return salesRoleKey_(user) === 'sales'; }
function isSalesManagerUser_(user) { return salesRoleKey_(user) === 'sales manager'; }
function isSalesScopedUser_(user) { return isSalesUser_(user) || isSalesManagerUser_(user); }

function salesNameKey_(v) { return String(v || '').replace(/\s+/g, ' ').trim().toLowerCase(); }

function salesRowIsShare_(r) {
  return String(r.dealStatus || '').trim().toLowerCase() === 'share' &&
    salesNameKey_(r.sharedWith) && salesNameKey_(r.sharedWith) !== salesNameKey_(r.salesName);
}

/* Sales name -> manager, from the deals themselves. */
function salesManagerMap_(rows) {
  const map = {};
  (rows || []).forEach(r => {
    const s = salesNameKey_(r.salesName), m = salesNameKey_(r.salesManager);
    if (s && m && !map[s]) map[s] = m;
  });
  return map;
}

function salesRowInScope_(r, user, managerMap) {
  const me = salesNameKey_(user.name);
  if (!me) return false;
  const primary = salesNameKey_(r.salesName);
  const shared = salesRowIsShare_(r) ? salesNameKey_(r.sharedWith) : '';
  if (primary === me || shared === me) return true;
  if (isSalesManagerUser_(user)) {
    if (salesNameKey_(r.salesManager) === me) return true;
    if (shared && managerMap[shared] === me) return true;
  }
  return false;
}

/* Narrows a full dashboard payload to the user's scope. */
function scopeDashboardForSales_(data, user) {
  const all = data.rows || [];
  const managerMap = salesManagerMap_(all);
  const rows = all.filter(r => salesRowInScope_(r, user, managerMap));

  // Keep only filter options that exist in the user's rows.
  const present = new Set();
  rows.forEach(r => Object.keys(r).forEach(k => {
    const v = r[k];
    if (v !== null && v !== undefined && v !== '') present.add(String(v).trim());
  }));
  const options = {};
  Object.keys(data.options || {}).forEach(k => {
    const v = data.options[k];
    options[k] = Array.isArray(v)
      ? v.filter(x => (x && typeof x === 'object') ? true : present.has(String(x).trim()))
      : v;
  });

  return Object.assign({}, data, {
    rows: rows,
    options: options,
    brokers: { sheetName: '', rows: [], options: {}, availableFields: {} },
    access: {
      role: isSalesManagerUser_(user) ? 'Sales Manager' : 'Sales',
      salesName: user.name,
      readOnly: false,
      allowedModules: ['transactions', 'inventory']
    }
  });
}

/* Targets a sales user may see: their own (and their team's for managers). */
function scopeTargetsForSales_(rows, user) {
  const me = salesNameKey_(user.name);
  let team = new Set([me]);
  if (isSalesManagerUser_(user)) {
    try {
      const data = dashCacheRead_('ALL') || buildDashboardData_({ role: 'Admin' });
      const map = salesManagerMap_(data.rows || []);
      Object.keys(map).forEach(s => { if (map[s] === me) team.add(s); });
    } catch (err) {}
  }
  return rows.filter(t => {
    const scope = String(t.scope || '').toLowerCase(), name = salesNameKey_(t.name);
    if (scope === 'sales') return team.has(name);
    if (scope === 'manager') return isSalesManagerUser_(user) && name === me;
    return false;
  });
}
