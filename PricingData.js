/************************************************
 * EDEN BI V4 — PricingData.js
 *
 * Change unit prices (price per m²) from the system, admins only.
 *  - Writes the new "Price Per M²" on the unit's row in its Inventory sheet
 *    (columns found by header name, the row is checked against the unit
 *    code before writing).
 *  - Total Price: if the cell is a formula (e.g. =L4*G4) it is left alone
 *    and recalculates by itself; if it is a typed number it moves by
 *    (new − old price per m²) × area, so outdoor parts stay as they were.
 *  - A price per m² that is itself a formula is never overwritten.
 *  - Every change goes to the "Price History" tab (main spreadsheet) and
 *    the Activity Log.
 ************************************************/

const PRICE_HIST_SHEET_ = 'Price History';
const PRICE_HIST_HEAD_ = ['Time', 'Project', 'Unit Code', 'Tab', 'Row', 'Old Price/m²', 'New Price/m²', 'Change %', 'Old Total', 'New Total', 'Currency', 'Reason', 'Changed By'];
const PRICE_MAX_BATCH_ = 400;

function priceAdmin_(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Changing prices is for admins only.');
  return user;
}

/* req: { project, reason, changes: [{ unitCode, sourceSheetId, sourceRow, meterPrice }] } */
function setUnitPrices(authToken, req) {
  const user = priceAdmin_(authToken);
  req = req || {};
  const key = inventoryKey_(req.project);
  const cfg = INVENTORY_SOURCES[key];
  if (!cfg) throw new Error('Inventory source is not configured for this project.');
  const changes = (req.changes || []).filter(c => c && c.unitCode && Number(c.meterPrice) > 0);
  if (!changes.length) throw new Error('No price changes to save.');
  if (changes.length > PRICE_MAX_BATCH_) throw new Error('Too many units at once (' + changes.length + '). Save at most ' + PRICE_MAX_BATCH_ + ' at a time.');
  const reason = String(req.reason || '').trim();
  const who = String(user.name || user.username || 'Admin');

  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const bySheet = {};
  changes.forEach(c => { (bySheet[String(c.sourceSheetId)] = bySheet[String(c.sourceSheetId)] || []).push(c); });
  const results = [], hist = [];

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    Object.keys(bySheet).forEach(sid => {
      const sheet = ss.getSheets().find(s => String(s.getSheetId()) === sid);
      const list = bySheet[sid];
      if (!sheet) { list.forEach(c => results.push({ unitCode: c.unitCode, ok: false, message: 'Tab not found.' })); return; }
      const lastCol = sheet.getLastColumn();
      const top = sheet.getRange(1, 1, Math.min(30, sheet.getLastRow()), lastCol).getDisplayValues();
      const hr = inventoryFindHeaderRow_(top, cfg);
      if (hr < 0) { list.forEach(c => results.push({ unitCode: c.unitCode, ok: false, message: 'Headers not found.' })); return; }
      const idx = inventoryIndex_(top[hr]);
      const colOf = aliases => { for (const a of aliases || []) { const p = idx[inventoryHeaderKey_(a)]; if (p !== undefined) return p + 1; } return 0; };
      const codeCol = colOf(cfg.headers.unitCode), meterCol = colOf(cfg.headers.meterPrice), totalCol = colOf(cfg.headers.totalPrice), areaCol = colOf(cfg.headers.area);
      if (!codeCol || !meterCol) { list.forEach(c => results.push({ unitCode: c.unitCode, ok: false, message: 'Code / Price per m² column not found.' })); return; }

      list.forEach(c => {
        const row = Number(c.sourceRow);
        try {
          if (!row || row <= hr + 1 || row > sheet.getLastRow()) throw new Error('Row not found.');
          const live = String(sheet.getRange(row, codeCol).getDisplayValue()).trim().toLowerCase();
          if (live !== String(c.unitCode).trim().toLowerCase()) throw new Error('Rows moved — refresh the inventory and try again.');
          const mCell = sheet.getRange(row, meterCol);
          if (mCell.getFormula()) throw new Error('Price per m² is a formula in the sheet — change it there.');
          const oldMeter = inventoryNumber_(mCell.getDisplayValue());
          const newMeter = Math.round(Number(c.meterPrice) * 100) / 100;
          const area = areaCol ? inventoryNumber_(sheet.getRange(row, areaCol).getDisplayValue()) : 0;
          let oldTotal = 0, tCell = null;
          if (totalCol) { tCell = sheet.getRange(row, totalCol); oldTotal = inventoryNumber_(tCell.getDisplayValue()); }
          if (Math.abs(oldMeter - newMeter) < 0.005) { results.push({ unitCode: c.unitCode, ok: true, unchanged: true }); return; }
          mCell.setValue(newMeter);
          if (tCell && !tCell.getFormula() && area > 0) tCell.setValue(Math.round(oldTotal ? oldTotal + (newMeter - oldMeter) * area : newMeter * area));
          results.push({ unitCode: c.unitCode, ok: true, oldMeter: oldMeter, newMeter: newMeter, area: area, oldTotal: oldTotal, row: row, tab: sheet.getName(), tCell: tCell });
        } catch (e) {
          results.push({ unitCode: c.unitCode, ok: false, message: e && e.message ? e.message : String(e) });
        }
      });
    });
    SpreadsheetApp.flush();
    results.forEach(r => {
      if (!r.ok || r.unchanged) return;
      const newTotal = r.tCell ? inventoryNumber_(r.tCell.getDisplayValue()) : (r.area ? Math.round(r.newMeter * r.area) : 0);
      r.newTotal = newTotal;
      delete r.tCell;
      hist.push([new Date(), cfg.project, r.unitCode, r.tab, r.row, r.oldMeter, r.newMeter,
        r.oldMeter ? Math.round((r.newMeter / r.oldMeter - 1) * 10000) / 100 : '', r.oldTotal || '', newTotal || '', cfg.currency, reason, who]);
    });
    results.forEach(r => { delete r.tCell; });
    if (hist.length) {
      const sh = instSheet_(PRICE_HIST_SHEET_, PRICE_HIST_HEAD_, true);
      sh.getRange(sh.getLastRow() + 1, 1, hist.length, PRICE_HIST_HEAD_.length).setValues(hist);
      sh.getRange(sh.getLastRow() - hist.length + 1, 1, hist.length, 1).setNumberFormat('yyyy-MM-dd HH:mm');
    }
  } finally { lock.releaseLock(); }

  const changed = results.filter(r => r.ok && !r.unchanged), failed = results.filter(r => !r.ok);
  if (changed.length) {
    auditLog_(user, 'Prices changed', { project: cfg.project, unit: changed.length === 1 ? changed[0].unitCode : changed.length + ' units',
      details: changed.slice(0, 40).map(r => r.unitCode + ': ' + r.oldMeter + ' → ' + r.newMeter).join(' · ') + (changed.length > 40 ? ' …' : '') + (reason ? ' · Reason: ' + reason : '') });
  }
  let inventory = null;
  try { inventory = inventoryDataCached_(key, true, false); } catch (e) { console.warn('setUnitPrices refresh: ' + e); }
  return { changed: changed.length, failed: failed, results: results, inventory: inventory };
}

/* Price changes of a project (newest first). */
function getPriceHistory(authToken, project) {
  priceAdmin_(authToken);
  const p = String(project || '').trim().toUpperCase();
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  return instRows_(instSheet_(PRICE_HIST_SHEET_, PRICE_HIST_HEAD_, false))
    .filter(r => !p || String(r[1]).trim().toUpperCase() === p)
    .map(r => ({ time: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm') : String(r[0]), project: String(r[1]), unitCode: String(r[2]), tab: String(r[3]),
      oldMeter: Number(r[5]) || 0, newMeter: Number(r[6]) || 0, pct: r[7] === '' ? null : Number(r[7]), oldTotal: Number(r[8]) || 0, newTotal: Number(r[9]) || 0,
      currency: String(r[10] || ''), reason: String(r[11] || ''), by: String(r[12] || '') }))
    .reverse().slice(0, 3000);
}
