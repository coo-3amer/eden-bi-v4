/************************************************
 * EDEN BI V4 — InventorySync.js
 *
 * Keeps the unit status in the Inventory sheet in step with its deal:
 *   Deal "Reservation"                         -> Inventory "Reservation"
 *   Part of DP / Total DP / Client Sigend / …  -> Inventory "Closed"
 *   Deal "Cancel" (and no other active deal)   -> Inventory "Available"
 *
 * Runs when a deal is added or updated from the system, and when the
 * Status cell is edited by hand in the Deals sheet (LiveSync trigger).
 * Columns are found by header names only. Every change is written as a
 * note on the Inventory cell. A sync problem never blocks saving a deal.
 ************************************************/

const INV_SYNC_PROJECTS_ = ['EDEN_WALK', 'CITY_CENTER_GLDANI'];

/* Deal status -> Inventory status ('' = leave the unit alone). */
function invSyncTarget_(dealStatus) {
  const s = String(dealStatus || '').trim().toLowerCase();
  if (!s) return '';
  if (dealCancelledStatus_(s)) return 'Available';
  if (s.indexOf('reserv') >= 0) return 'Reservation';
  return 'Closed';
}

/* Never throws. Returns { changed, message }. */
function syncInventoryForDeal_(project, unitCode, dealStatus, opts) {
  opts = opts || {};
  const skip = msg => ({ changed: false, message: msg });
  try {
    const key = inventoryKey_(dealText_(project));
    if (INV_SYNC_PROJECTS_.indexOf(key) < 0) return skip('');
    const cfg = INVENTORY_SOURCES[key];
    const code = dealText_(unitCode).toLowerCase();
    const target = invSyncTarget_(dealStatus);
    if (!code || !target) return skip('');

    // Cancel: the unit is freed only if no other active deal holds it.
    if (target === 'Available') {
      const ds = getMainDealsSheet_();
      if (findExistingActiveMainDeal_(ds, mainDealsColumns_(ds), project, unitCode)) {
        return skip('Inventory not changed: another active deal holds ' + unitCode + '.');
      }
    }

    // Find the unit (from the cached inventory, then confirmed on the live sheet).
    const data = inventoryDataCached_(key, false, false);
    let matches = (data.rows || []).filter(r => String(r.unitCode || '').trim().toLowerCase() === code);
    const cat = dealText_(opts.inventoryCategory).toLowerCase();
    if (cat && matches.length > 1) {
      const byCat = matches.filter(r => String(r.inventoryCategory || '').toLowerCase() === cat);
      if (byCat.length) matches = byCat;
    }
    if (matches.length > 1) {
      const real = matches.filter(r => !/^(divided|share)$/i.test(String(r.status || '').trim()));
      if (real.length) matches = real;
    }
    if (matches.length !== 1) {
      return skip(matches.length ? 'Inventory not changed: ' + unitCode + ' appears more than once.' : 'Inventory not changed: ' + unitCode + ' was not found in the inventory.');
    }
    const m = matches[0];

    const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
    const sheet = ss.getSheets().find(s => Number(s.getSheetId()) === Number(m.sourceSheetId));
    if (!sheet) return skip('Inventory not changed: tab not found.');
    const lastCol = sheet.getLastColumn();
    const top = sheet.getRange(1, 1, Math.min(30, sheet.getLastRow()), lastCol).getDisplayValues();
    const hr = inventoryFindHeaderRow_(top, cfg);
    if (hr < 0) return skip('Inventory not changed: headers not found.');
    const idx = inventoryIndex_(top[hr]);
    const colOf = aliases => { for (const a of aliases || []) { const p = idx[inventoryHeaderKey_(a)]; if (p !== undefined) return p + 1; } return 0; };
    const statusCol = colOf(cfg.headers.status || ['Status']);
    const codeCol = colOf(cfg.headers.unitCode);
    if (!statusCol || !codeCol) return skip('Inventory not changed: Status / Code column not found.');

    // Safety: the row must still be this unit.
    const liveCode = String(sheet.getRange(m.sourceRow, codeCol).getDisplayValue()).trim().toLowerCase();
    if (liveCode !== code) return skip('Inventory not changed: rows moved, refresh the inventory.');
    const cell = sheet.getRange(m.sourceRow, statusCol);
    const current = String(cell.getDisplayValue()).trim();
    const cur = current.toLowerCase();

    // Rules that protect manual decisions.
    if (cur === target.toLowerCase()) return skip('');
    if (target === 'Closed' && ['closed', 'done', 'sold', 'delivered'].indexOf(cur) >= 0) return skip('');
    if (target === 'Reservation' && ['closed', 'done', 'sold', 'delivered'].indexOf(cur) >= 0) return skip('');
    if (target === 'Available' && !(cur === 'closed' || cur.indexOf('reserv') >= 0)) return skip('');

    // Use the exact spelling of the cell's dropdown.
    let value = target;
    const rule = cell.getDataValidation();
    if (rule) {
      let list = [];
      const t = rule.getCriteriaType(), c = rule.getCriteriaValues();
      if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) list = c[0] || [];
      else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && c[0]) list = c[0].getDisplayValues().flat();
      if (list.length) {
        const hit = list.find(x => String(x).trim().toLowerCase() === target.toLowerCase());
        if (hit) value = String(hit).trim();
        else if (!rule.getAllowInvalid()) return skip('Inventory not changed: "' + target + '" is not in the Status list.');
      }
    }

    cell.setValue(value);
    try {
      const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
      const prev = cell.getNote();
      cell.setNote((prev ? prev + '\n' : '') + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm') +
        ' • EDEN BI' + (opts.by ? ' (' + opts.by + ')' : '') + ': ' + (current || '—') + ' → ' + value +
        ' (deal ' + dealText_(dealStatus) + ')');
    } catch (e) {}

    // Keep the cached inventory current (script edits don't fire LiveSync).
    try {
      m.status = value;
      const counts = {};
      data.rows.forEach(r => counts[r.status] = (counts[r.status] || 0) + 1);
      const low = r => String(r.status || '').trim().toLowerCase();
      if (data.stats) {
        data.stats.statusCounts = counts;
        data.stats.available = data.rows.filter(r => low(r) === 'available').length;
        data.stats.reservation = data.rows.filter(r => low(r).indexOf('reservation') >= 0).length;
        data.stats.closed = data.rows.filter(r => ['closed', 'sold'].indexOf(low(r)) >= 0).length;
      }
      if (data.options && data.options.statuses && data.options.statuses.indexOf(value) < 0) data.options.statuses.push(value);
      data.generatedAt = new Date().toISOString();
      dashCacheWrite_('INV_' + key, data);
    } catch (e) { console.warn('Inventory cache update: ' + e); }

    return { changed: true, from: current, to: value, message: 'Inventory: ' + m.unitCode + ' ' + (current || '—') + ' → ' + value + '.' };
  } catch (err) {
    console.warn('syncInventoryForDeal_: ' + err);
    return skip('Inventory not changed: ' + (err && err.message ? err.message : err));
  }
}

/* Hand edits of the Status column in the Deals sheet (called by LiveSync). */
function syncInventoryFromDealsEdit_(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getSheetId() !== DEALS_GID) return;
    const cols = mainDealsColumns_(sheet);
    const sc = mainDealsCol_(cols, ['deal status|status', 'Status']);
    const r0 = e.range.getRow(), c0 = e.range.getColumn();
    const nR = e.range.getNumRows(), nC = e.range.getNumColumns();
    if (!sc || sc < c0 || sc > c0 + nC - 1) return;
    const first = Math.max(r0, DATA_START_ROW), last = Math.min(r0 + nR - 1, first + 49);
    if (last < first) return;
    const pc = mainDealsCol_(cols, ['Project']), uc = mainDealsCol_(cols, ['Unit Code']);
    if (!pc || !uc) return;
    const row = (c) => sheet.getRange(first, c, last - first + 1, 1).getDisplayValues().map(x => x[0]);
    const P = row(pc), U = row(uc), S = row(sc);
    let by = '';
    try { by = (e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (err) {}
    for (let i = 0; i < P.length; i++) {
      if (P[i] && U[i] && S[i]) syncInventoryForDeal_(P[i], U[i], S[i], { by: by ? 'sheet edit by ' + by : 'sheet edit' });
    }
  } catch (err) { console.warn('syncInventoryFromDealsEdit_: ' + err); }
}
