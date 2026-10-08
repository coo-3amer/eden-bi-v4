/************************************************
 * EDEN BI V4 — FloorMaps.js
 *
 * Floor Map page: a floor plan image with every unit pinned on it,
 * coloured by its live Inventory status.
 * Pins (unit code -> x/y on the image, 0..1) are kept in the
 * "Floor Map Pins" tab of the main spreadsheet. Admins place them
 * once from the page; everyone else just views.
 ************************************************/

const FLOOR_MAPS_ = {
  EW_G: { key: 'EW_G', project: 'EDEN WALK', inventory: 'EDEN_WALK', title: 'Eden Walk — Ground Floor',
          image: 'maps/eden-walk-ground.jpg', floor: /ground/i }
};
const FLOOR_PINS_SHEET_ = 'Floor Map Pins';

function floorMapIsAdmin_(user) {
  const role = String(user && user.role || '').trim().toLowerCase();
  return role === 'admin' || role === 'super admin';
}

function floorPinsSheet_(create) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(FLOOR_PINS_SHEET_);
  if (!sh && create) {
    sh = ss.insertSheet(FLOOR_PINS_SHEET_);
    sh.getRange(1, 1, 1, 6).setValues([['Map', 'Unit Code', 'X', 'Y', 'Updated By', 'Updated At']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function floorPinsRead_(mapKey) {
  const sh = floorPinsSheet_(false);
  const pins = {};
  if (!sh || sh.getLastRow() < 2) return pins;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
    if (String(r[0]) !== mapKey || !String(r[1]).trim()) return;
    const x = Number(r[2]), y = Number(r[3]);
    if (isFinite(x) && isFinite(y)) pins[String(r[1]).trim()] = { x: x, y: y };
  });
  return pins;
}

/* Units of the floor (from the cached Inventory) + saved pins. */
function getFloorMap(authToken, mapKey) {
  const user = validateAuthToken_(authToken);
  const cfg = FLOOR_MAPS_[String(mapKey || 'EW_G')];
  if (!cfg) throw new Error('This floor map is not set up.');
  const inv = inventoryDataCached_(cfg.inventory, false, false);
  const units = (inv.rows || [])
    .filter(r => cfg.floor.test(String(r.floor || '')) && String(r.unitCode || '').trim())
    .map(r => ({
      code: String(r.unitCode).trim(),
      category: r.inventoryCategory || '',
      status: r.status || '',
      type: r.type || r.unitType || '',
      area: r.area || 0,
      outdoor: r.outdoorAreaNumber || 0,
      meterPrice: r.meterPrice || 0,
      totalPrice: r.totalPrice || 0,
      finishing: r.finishingType || '',
      phase: r.phase || ''
    }));
  return {
    map: { key: cfg.key, title: cfg.title, image: cfg.image, project: cfg.project, currency: inv.currency || 'EGP' },
    units: units,
    pins: floorPinsRead_(cfg.key),
    canEdit: floorMapIsAdmin_(user),
    generatedAt: inv.generatedAt || ''
  };
}

/* Admin: place / move a pin (x, y between 0 and 1) or remove it (x = null). */
function saveFloorMapPin(authToken, mapKey, unitCode, x, y) {
  const user = validateAuthToken_(authToken);
  if (!floorMapIsAdmin_(user)) throw new Error('Only admins can place units on the map.');
  const cfg = FLOOR_MAPS_[String(mapKey || '')];
  if (!cfg) throw new Error('This floor map is not set up.');
  const code = String(unitCode || '').trim();
  if (!code) throw new Error('Pick a unit first.');
  const remove = x === null || x === '' || x === undefined;
  const nx = Number(x), ny = Number(y);
  if (!remove && !(nx >= 0 && nx <= 1 && ny >= 0 && ny <= 1)) throw new Error('The point is outside the map.');

  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { throw new Error('Busy, try again.'); }
  try {
    const sh = floorPinsSheet_(true);
    const last = sh.getLastRow();
    let row = 0;
    if (last >= 2) {
      const keys = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === cfg.key && String(keys[i][1]).trim().toLowerCase() === code.toLowerCase()) { row = i + 2; break; }
      }
    }
    if (remove) {
      if (row) sh.deleteRow(row);
    } else {
      const vals = [[cfg.key, code, Math.round(nx * 10000) / 10000, Math.round(ny * 10000) / 10000, user.name || user.username, new Date()]];
      if (row) sh.getRange(row, 1, 1, 6).setValues(vals);
      else sh.appendRow(vals[0]);
    }
    return { success: true, unitCode: code, removed: remove };
  } finally {
    lock.releaseLock();
  }
}
