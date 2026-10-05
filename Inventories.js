const INVENTORY_SOURCES = {
  KOBULETI: {
    project: 'KOBULETI',
    title: 'Kobuleti',
    spreadsheetId: '1QnBdWz2xDvZ23J_J-rkz_G6Ko1_5QAKKY3UQ9oCoUPs',
    sheetName: 'B',
    currency: 'USD',
    companyMode: 'KOBULETI_MARKER',
    headers: {
      unitCode: ['Unit Code'],
      floor: ['Unit/Floor', 'Floor'],
      unitType: ['Unit Type'],
      type: ['Type'],
      area: ['Total gross', 'Area'],
      view: ['View'],
      meterPrice: ['Price Per M² (White Frame)', 'Price Per M²', 'Price Per M2', 'Meter Price'],
      totalPrice: ['Total price (White Frame)', 'Total Price', 'Price'],
      status: ['Status']
    }
  },

  CITY_CENTER_GLDANI: {
    project: 'CITY CENTER GLDANI',
    title: 'City Center Gldani',
    spreadsheetId: '166g24g7Tk0Z_5x9bj8yabQoamoBVZ6TIk7ruzzd8w1Y',
    sheetId: 578242077,
    currency: 'USD',
    companyMode: 'EDEN',
    headers: {
      unitCode: ['Unite Code', 'Unit Code', 'Code'],
      floor: ['Unit/Floor', 'Floor'],
      unitType: ['Model'],
      type: ['Type'],
      area: ['Area'],
      view: ['View'],
      meterPrice: ['Price Per M²', 'Price Per M2', 'Price/m²'],
      totalPrice: ['Total Price'],
      status: ['Status'],
      block: ['Block'],
      downPayment: ['Down Payment'],
      installments: ['Installments']
    }
  },

  EDEN_WALK: {
    project: 'EDEN WALK',
    title: 'Eden Walk',
    spreadsheetId: '1UTOJmajKptMNdZK9aQ6-sgOAtYzY4NBgHhuf91-gPaI',
    // Eden Walk now reads all live inventory tabs by TAB NAME.
    // No column letters/indexes are used for field mapping.
    sheetNames: ['AV', 'Shared Units', 'Divided Units'],
    sheetCategories: {
      'AV': 'Main Unit',
      'Shared Units': 'Shared Unit',
      'Divided Units': 'Divided Unit'
    },
    primarySheetName: 'AV',
    currency: 'EGP',
    companyMode: 'EDEN',
    headers: {
      unitCode: ['Code', 'Unit Code'],
      floor: ['Unit/Floor', 'Floor'],
      unitType: ['Type'],
      type: ['Type'],
      area: ['Indoor', 'Area'],
      outdoorArea: ['Outdoor'],
      view: ['View'],
      meterPrice: ['Indoor Price/m', 'Price Per M²', 'Price Per M2'],
      outdoorMeterPrice: ['Outdoor Price/m'],
      totalPrice: ['Total Price'],
      status: ['Status'],
      finishingType: ['Finishing type', 'Finishing Type'],
      phase: ['Phase']
    }
  }
};

function inventoryKey_(project) {
  const raw = String(project || 'KOBULETI').trim().toUpperCase();
  const aliases = {
    'KOBULETI': 'KOBULETI',
    'CITY CENTER GLDANI': 'CITY_CENTER_GLDANI',
    'CITY CENTRE GLDANI': 'CITY_CENTER_GLDANI',
    'CCG': 'CITY_CENTER_GLDANI',
    'CITY_CENTER_GLDANI': 'CITY_CENTER_GLDANI',
    'EDEN WALK': 'EDEN_WALK',
    'EDEN_WALK': 'EDEN_WALK'
  };
  return aliases[raw] || raw;
}

function inventorySheet_(ss, cfg) {
  if (cfg.sheetName) {
    const byName = ss.getSheetByName(cfg.sheetName);
    if (!byName) throw new Error('Inventory tab "' + cfg.sheetName + '" was not found.');
    return byName;
  }

  if (cfg.sheetId !== undefined && cfg.sheetId !== null) {
    const byId = ss.getSheets().find(s => Number(s.getSheetId()) === Number(cfg.sheetId));
    if (!byId) throw new Error('The inventory tab linked by gid=' + cfg.sheetId + ' was not found.');
    return byId;
  }

  throw new Error('Inventory tab is not configured.');
}


function inventorySheets_(ss, cfg) {
  if (Array.isArray(cfg.sheetNames) && cfg.sheetNames.length) {
    return cfg.sheetNames.map(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error('Inventory tab "' + name + '" was not found.');
      return sheet;
    });
  }
  return [inventorySheet_(ss, cfg)];
}

function inventoryCategoryForSheet_(cfg, sheet) {
  const name = sheet.getName();
  return (cfg.sheetCategories && cfg.sheetCategories[name]) || 'Main Unit';
}

function inventoryReadSheetRows_(sheet, cfg, key) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];

  const headerRow = inventoryFindHeaderRow_(values, cfg);
  if (headerRow === -1) {
    throw new Error('Inventory headers were not found in tab "' + sheet.getName() + '".');
  }

  const headers = values[headerRow].map(inventoryClean_);
  const index = inventoryIndex_(headers);
  const rows = [];
  let currentFloor = '';
  const inventoryCategory = inventoryCategoryForSheet_(cfg, sheet);

  values.slice(headerRow + 1).forEach((row, offset) => {
    const unitCode = inventoryGet_(row, index, cfg.headers.unitCode);
    const status = inventoryGet_(row, index, cfg.headers.status);
    const floorValue = inventoryGet_(row, index, cfg.headers.floor);

    if (floorValue && !unitCode) currentFloor = floorValue;
    if (!unitCode) return;
    if (floorValue) currentFloor = floorValue;

    const areaText = inventoryGet_(row, index, cfg.headers.area);
    const outdoorAreaText = inventoryGet_(row, index, cfg.headers.outdoorArea);
    const meterPriceText = inventoryGet_(row, index, cfg.headers.meterPrice);
    const outdoorMeterPriceText = inventoryGet_(row, index, cfg.headers.outdoorMeterPrice);
    const totalPriceText = inventoryGet_(row, index, cfg.headers.totalPrice);
    const company = inventoryCompany_(cfg, row, status);

    rows.push({
      project: cfg.project,
      projectKey: key,
      sourceRow: headerRow + 2 + offset,
      sourceSheetName: sheet.getName(),
      sourceSheetId: sheet.getSheetId(),
      inventoryCategory: inventoryCategory,
      unitKey: sheet.getName() + '::' + unitCode,
      floor: floorValue || currentFloor,
      unitType: inventoryGet_(row, index, cfg.headers.unitType),
      unitCode: unitCode,
      type: inventoryGet_(row, index, cfg.headers.type),
      block: inventoryGet_(row, index, cfg.headers.block),
      totalGross: areaText,
      area: inventoryNumber_(areaText),
      outdoorArea: outdoorAreaText,
      outdoorAreaNumber: inventoryNumber_(outdoorAreaText),
      view: inventoryGet_(row, index, cfg.headers.view),
      meterPrice: inventoryNumber_(meterPriceText),
      outdoorMeterPrice: inventoryNumber_(outdoorMeterPriceText),
      totalPrice: inventoryNumber_(totalPriceText),
      currency: cfg.currency,
      status: status || 'Not Specified',
      company: company,
      finishingType: inventoryGet_(row, index, cfg.headers.finishingType),
      phase: inventoryGet_(row, index, cfg.headers.phase),
      downPayment: inventoryGet_(row, index, cfg.headers.downPayment),
      installments: inventoryGet_(row, index, cfg.headers.installments)
    });
  });

  return rows;
}

function inventoryClean_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
}

function inventoryHeaderKey_(value) {
  return inventoryClean_(value)
    .toLowerCase()
    .replace(/²/g, '2')
    .replace(/[^\p{L}\p{N}]+/gu, '');
}

function inventoryNumber_(value) {
  let text = String(value == null ? '' : value).trim();
  if (!text) return 0;

  // Keep the first valid numeric token and remove currency/thousand separators.
  const negative = /^\s*-/.test(text);
  text = text.replace(/,/g, '').replace(/[^\d.]/g, '');
  const number = Number(text);
  if (!Number.isFinite(number)) return 0;
  return negative ? -number : number;
}

function inventoryFindHeaderRow_(values, cfg) {
  const unitAliases = (cfg.headers.unitCode || []).map(inventoryHeaderKey_);
  const statusAliases = (cfg.headers.status || ['Status']).map(inventoryHeaderKey_);

  for (let r = 0; r < Math.min(values.length, 30); r++) {
    const rowKeys = values[r].map(inventoryHeaderKey_);
    const hasUnit = unitAliases.some(alias => rowKeys.includes(alias));
    const hasStatus = statusAliases.some(alias => rowKeys.includes(alias));
    if (hasUnit && hasStatus) return r;
  }
  return -1;
}

function inventoryIndex_(headers) {
  const index = {};
  headers.forEach((h, i) => {
    const key = inventoryHeaderKey_(h);
    if (key && index[key] === undefined) index[key] = i;
  });
  return index;
}

function inventoryGet_(row, index, aliases) {
  for (const alias of aliases || []) {
    const pos = index[inventoryHeaderKey_(alias)];
    if (pos !== undefined) return inventoryClean_(row[pos]);
  }
  return '';
}

function inventoryCompany_(cfg, row, status) {
  if (cfg.companyMode === 'EDEN') return 'EDEN';

  // Kobuleti-specific ownership rule:
  // J = W => LOFT
  // J blank + Reservation/Sold/Closed => EDEN
  const marker = inventoryClean_(row[9]).toUpperCase();
  const state = inventoryClean_(status).toLowerCase();

  if (marker === 'W') return 'LOFT';
  if (!marker && (state.includes('reservation') || state === 'sold' || state === 'closed')) return 'EDEN';
  return '';
}

function getInventoryData(authToken, project, forceRefresh) {
  validateAuthToken_(authToken);
  return inventoryDataCached_(project, forceRefresh, !!forceRefresh);
}

const INVENTORY_CACHE_MAX_AGE_MS_ = 6 * 60 * 1000;   // refreshed every 5 min by warmDashboardCache

/* Inventory of one project. Kept compressed in the cache (any size), so
   unit lists and unit checks answer in about a second. */
function inventoryDataCached_(project, forceRefresh, trackMovements) {
  const key = inventoryKey_(project);
  const cfg = INVENTORY_SOURCES[key];
  if (!cfg) throw new Error('Inventory source is not configured for this project.');

  const scope = 'INV_' + key;
  if (!forceRefresh) {
    const cached = dashCacheRead_(scope, INVENTORY_CACHE_MAX_AGE_MS_);
    if (cached) return cached;
  }

  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheets = inventorySheets_(ss, cfg);
  const rows = sheets.reduce((allRows, sheet) => {
    return allRows.concat(inventoryReadSheetRows_(sheet, cfg, key));
  }, []);

  // On an explicit refresh, compare against the persistent snapshot and log status changes.
  if (trackMovements && key === 'EDEN_WALK') inventoryTrackMovements_(cfg, key, rows, ss);

  const unique = field => [...new Set(rows.map(r => inventoryClean_(r[field])).filter(Boolean))]
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));

  const statusCounts = {};
  rows.forEach(r => statusCounts[r.status] = (statusCounts[r.status] || 0) + 1);

  const result = {
    project: cfg.project,
    projectKey: key,
    title: cfg.title,
    currency: cfg.currency,
    sheetName: sheets.map(s => s.getName()).join(' + '),
    sheetNames: sheets.map(s => s.getName()),
    sheetId: sheets[0] ? sheets[0].getSheetId() : null,
    spreadsheetUrl: ss.getUrl() + (sheets[0] ? '#gid=' + sheets[0].getSheetId() : ''),
    generatedAt: new Date().toISOString(),
    rows: rows,
    options: {
      floors: unique('floor'),
      unitTypes: unique('unitType'),
      types: unique('type'),
      statuses: unique('status'),
      companies: unique('company'),
      inventoryCategories: unique('inventoryCategory')
    },
    stats: {
      totalUnits: rows.length,
      available: rows.filter(r => inventoryClean_(r.status).toLowerCase() === 'available').length,
      reservation: rows.filter(r => inventoryClean_(r.status).toLowerCase().includes('reservation')).length,
      closed: rows.filter(r => ['closed', 'sold'].includes(inventoryClean_(r.status).toLowerCase())).length,
      totalValue: rows.reduce((sum, r) => sum + Number(r.totalPrice || 0), 0),
      statusCounts: statusCounts
    }
  };

  dashCacheWrite_(scope, result);
  return result;
}

function getInventoryUnit(authToken, project, unitCode, inventoryCategory) {
  const data = getInventoryData(authToken, project, false);
  const code = inventoryClean_(unitCode).toLowerCase();
  const category = inventoryClean_(inventoryCategory).toLowerCase();

  const matches = data.rows.filter(r => {
    const sameCode = inventoryClean_(r.unitCode).toLowerCase() === code;
    const sameCategory = !category || inventoryClean_(r.inventoryCategory).toLowerCase() === category;
    return sameCode && sameCategory;
  });

  if (!matches.length) return null;
  if (matches.length === 1) return matches[0];

  // If the caller did not specify a category and the same unit code exists
  // in more than one Eden Walk tab, do not silently validate the wrong record.
  if (!category) {
    return {
      ambiguous: true,
      unitCode: unitCode,
      matches: matches.map(r => ({
        inventoryCategory: r.inventoryCategory,
        sourceSheetName: r.sourceSheetName,
        status: r.status,
        area: r.area,
        totalPrice: r.totalPrice
      }))
    };
  }

  return matches[0];
}

function validateInventoryUnitForDeal(authToken, project, unitCode, inventoryCategory) {
  validateAuthToken_(authToken);
  const unit = getInventoryUnit(authToken, project, unitCode, inventoryCategory);

  if (!unit) {
    return { allowed: false, reason: 'UNIT_NOT_FOUND', message: 'Unit was not found in the inventory.' };
  }

  if (unit.ambiguous) {
    return {
      allowed: false,
      reason: 'AMBIGUOUS_UNIT_CODE',
      message: 'This unit code exists in more than one Eden Walk inventory category. Select the exact Unit Category first.',
      matches: unit.matches
    };
  }

  const status = inventoryClean_(unit.status).toLowerCase();

  if (status === 'available') {
    return { allowed: true, status: unit.status, unit: unit, message: 'Unit is available.' };
  }

  if (status.includes('reservation')) {
    return { allowed: false, reason: 'RESERVED', status: unit.status, unit: unit, message: 'This unit is currently reserved.' };
  }

  if (status === 'closed' || status === 'sold' || status === 'delivered') {
    return { allowed: false, reason: 'SOLD', status: unit.status, unit: unit, message: 'This unit is already sold/closed and cannot be added to a new deal.' };
  }

  if (status.includes('not available')) {
    return { allowed: false, reason: 'NOT_AVAILABLE', status: unit.status, unit: unit, message: 'This unit is not available.' };
  }

  return {
    allowed: false,
    reason: 'STATUS_REVIEW',
    status: unit.status,
    unit: unit,
    message: 'Unit status requires review before adding a deal.'
  };
}

// ===== INVENTORY MOVEMENT HISTORY V1 =====
const INVENTORY_MOVEMENT_LOG_SHEET_ = 'BI Inventory Movements';
const INVENTORY_SNAPSHOT_SHEET_ = 'BI Inventory Snapshot';

function inventoryMovementKey_(r) {
  return [r.projectKey || r.project, r.sourceSheetName || '', r.inventoryCategory || '', r.unitCode || '']
    .map(inventoryClean_).join('||').toLowerCase();
}

function inventoryMovementStatusKey_(value) {
  return inventoryClean_(value).toLowerCase();
}

function inventoryEnsureTrackingSheets_(ss) {
  let log = ss.getSheetByName(INVENTORY_MOVEMENT_LOG_SHEET_);
  if (!log) log = ss.insertSheet(INVENTORY_MOVEMENT_LOG_SHEET_);
  const logHeaders = ['Timestamp','Date','Project','Unit Code','Unit Category','Source Tab','Previous Status','New Status','Floor','Unit Type','Type','Area','Total Price','Currency','Movement Type'];
  if (log.getLastRow() === 0) {
    log.getRange(1,1,1,logHeaders.length).setValues([logHeaders]);
  } else if (inventoryClean_(log.getRange(1,15).getDisplayValue()) !== 'Movement Type') {
    log.getRange(1,15).setValue('Movement Type');
  }

  let snap = ss.getSheetByName(INVENTORY_SNAPSHOT_SHEET_);
  if (!snap) snap = ss.insertSheet(INVENTORY_SNAPSHOT_SHEET_);
  const snapHeaders = ['Key','Project','Unit Code','Unit Category','Source Tab','Status','Floor','Unit Type','Type','Area','Total Price','Currency','Updated At'];
  if (snap.getLastRow() === 0) snap.getRange(1,1,1,snapHeaders.length).setValues([snapHeaders]);
  return {log:log, snapshot:snap};
}

function inventorySeedCurrentNonAvailable_(cfg, rows, tracking, now, dateText) {
  // Idempotent catch-up:
  // every current non-available unit is added once if that exact detected status
  // has never been logged before. This fixes upgrades where an older seed flag
  // was already stored before the current logging logic was installed.
  const existing = new Set();

  if (tracking.log.getLastRow() > 1) {
    const vals = tracking.log.getRange(2,1,tracking.log.getLastRow()-1,15).getDisplayValues();
    vals.forEach(v => {
      const sig = [v[3], v[4], v[5], v[7]]
        .map(inventoryClean_)
        .join('||')
        .toLowerCase();
      if (sig) existing.add(sig);
    });
  }

  const initial = [];
  rows.forEach(r => {
    const status = inventoryMovementStatusKey_(r.status);
    if (!status || status === 'available') return;

    const sig = [r.unitCode, r.inventoryCategory, r.sourceSheetName, r.status]
      .map(inventoryClean_)
      .join('||')
      .toLowerCase();

    if (existing.has(sig)) return;

    initial.push([
      now, dateText, cfg.project, r.unitCode || '', r.inventoryCategory || '', r.sourceSheetName || '',
      '', r.status || '', r.floor || '', r.unitType || '', r.type || '', Number(r.area || 0),
      Number(r.totalPrice || 0), r.currency || cfg.currency || '', 'INITIAL DETECTED'
    ]);
    existing.add(sig);
  });

  if (initial.length) {
    tracking.log.getRange(tracking.log.getLastRow()+1,1,initial.length,15).setValues(initial);
  }

  return initial.length;
}

function inventoryTrackMovements_(cfg, key, rows, ss) {
  // Eden Walk first. Architecture can be extended to the other inventory sources later.
  if (key !== 'EDEN_WALK') return {logged:0, initialized:false};
  const tracking = inventoryEnsureTrackingSheets_(ss);
  const snap = tracking.snapshot;
  const oldValues = snap.getLastRow() > 1 ? snap.getRange(2,1,snap.getLastRow()-1,13).getDisplayValues() : [];
  const oldMap = {};
  oldValues.forEach(r => { if (r[0]) oldMap[String(r[0]).toLowerCase()] = r; });

  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const dateText = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const initialDetected = inventorySeedCurrentNonAvailable_(cfg, rows, tracking, now, dateText);
  const movements = [];

  rows.forEach(r => {
    const k = inventoryMovementKey_(r);
    const old = oldMap[k];
    if (!old) return; // New units are not treated as status movements.
    const previousStatus = old[5] || '';
    const newStatus = r.status || '';
    if (inventoryMovementStatusKey_(previousStatus) === inventoryMovementStatusKey_(newStatus)) return;
    movements.push([
      now, dateText, cfg.project, r.unitCode || '', r.inventoryCategory || '', r.sourceSheetName || '',
      previousStatus, newStatus, r.floor || '', r.unitType || '', r.type || '', Number(r.area || 0),
      Number(r.totalPrice || 0), r.currency || cfg.currency || '', 'STATUS CHANGE'
    ]);
  });

  if (movements.length) {
    tracking.log.getRange(tracking.log.getLastRow()+1,1,movements.length,movements[0].length).setValues(movements);
  }

  const snapshotRows = rows.map(r => [
    inventoryMovementKey_(r), cfg.project, r.unitCode || '', r.inventoryCategory || '', r.sourceSheetName || '',
    r.status || '', r.floor || '', r.unitType || '', r.type || '', Number(r.area || 0), Number(r.totalPrice || 0),
    r.currency || cfg.currency || '', now
  ]);
  if (snap.getLastRow() > 1) snap.getRange(2,1,snap.getLastRow()-1,13).clearContent();
  if (snapshotRows.length) snap.getRange(2,1,snapshotRows.length,13).setValues(snapshotRows);
  return {logged:movements.length, initialDetected:initialDetected, initialized:oldValues.length===0};
}

function inventoryMovementDateText_(value, tz) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }

  const raw = inventoryClean_(value);
  if (!raw) return '';

  const iso = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    return iso[1] + '-' + String(iso[2]).padStart(2,'0') + '-' + String(iso[3]).padStart(2,'0');
  }

  const us = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (us) {
    return us[3] + '-' + String(us[1]).padStart(2,'0') + '-' + String(us[2]).padStart(2,'0');
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  }

  return raw;
}

function getInventoryMovements(authToken, project, startDate, endDate) {
  validateAuthToken_(authToken);

  const key = inventoryKey_(project || 'EDEN WALK');
  const cfg = INVENTORY_SOURCES[key];
  if (!cfg) throw new Error('Inventory source is not configured for this project.');

  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);

  if (key === 'EDEN_WALK') {
    const sourceSheets = inventorySheets_(ss, cfg);
    const currentRows = sourceSheets.reduce(
      (all, sheet) => all.concat(inventoryReadSheetRows_(sheet, cfg, key)), []
    );
    inventoryTrackMovements_(cfg, key, currentRows, ss);
  }

  const log = ss.getSheetByName(INVENTORY_MOVEMENT_LOG_SHEET_);

  const emptyResult = {
    project: cfg.project,
    from: startDate || '',
    to: endDate || startDate || '',
    rows: [],
    stats: {
      reservation: 0,
      closed: 0,
      returnedAvailable: 0,
      reservationValue: 0,
      closedValue: 0
    },
    groups: {
      reservations: [],
      closed: [],
      returnedAvailable: []
    }
  };

  if (!log || log.getLastRow() <= 1) return emptyResult;

  const values = log.getDataRange().getDisplayValues();
  if (values.length <= 1) return emptyResult;

  const headers = values[0].map(inventoryClean_);
  const index = inventoryIndex_(headers);
  const get_ = (row, header) => inventoryGet_(row, index, [header]);

  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const from = inventoryMovementDateText_(startDate, tz) || today;
  const to = inventoryMovementDateText_(endDate, tz) || from;

  const rows = values.slice(1)
    .map(row => ({
      timestamp: get_(row, 'Timestamp'),
      date: inventoryMovementDateText_(get_(row, 'Date'), tz),
      project: get_(row, 'Project'),
      unitCode: get_(row, 'Unit Code'),
      inventoryCategory: get_(row, 'Unit Category'),
      sourceSheetName: get_(row, 'Source Tab'),
      previousStatus: get_(row, 'Previous Status'),
      newStatus: get_(row, 'New Status'),
      floor: get_(row, 'Floor'),
      unitType: get_(row, 'Unit Type'),
      type: get_(row, 'Type'),
      area: inventoryNumber_(get_(row, 'Area')),
      totalPrice: inventoryNumber_(get_(row, 'Total Price')),
      currency: get_(row, 'Currency') || cfg.currency || '',
      movementType: get_(row, 'Movement Type') || 'STATUS CHANGE'
    }))
    .filter(r =>
      inventoryClean_(r.project).toUpperCase() === inventoryClean_(cfg.project).toUpperCase() &&
      r.date &&
      r.date >= from &&
      r.date <= to &&
      inventoryClean_(r.movementType).toUpperCase() === 'STATUS CHANGE'
    )
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));

  const statusKey_ = value => inventoryClean_(value).toLowerCase();
  const isReservation_ = r => statusKey_(r.newStatus).includes('reservation');
  const isClosed_ = r => ['closed', 'sold'].includes(statusKey_(r.newStatus));
  const isReturned_ = r =>
    statusKey_(r.newStatus) === 'available' &&
    statusKey_(r.previousStatus) !== 'available';

  function uniqueByUnit_(inputRows) {
    const map = new Map();

    inputRows.forEach(r => {
      const unitKey = [
        inventoryClean_(r.sourceSheetName),
        inventoryClean_(r.inventoryCategory),
        inventoryClean_(r.unitCode)
      ].join('||').toLowerCase();

      if (!map.has(unitKey)) map.set(unitKey, r);
    });

    return [...map.values()].sort((a, b) =>
      String(a.unitCode).localeCompare(String(b.unitCode), undefined, { numeric: true })
    );
  }

  const reservations = uniqueByUnit_(rows.filter(isReservation_));
  const closed = uniqueByUnit_(rows.filter(isClosed_));
  const returnedAvailable = uniqueByUnit_(rows.filter(isReturned_));

  return {
    project: cfg.project,
    from,
    to,
    rows,
    stats: {
      reservation: reservations.length,
      closed: closed.length,
      returnedAvailable: returnedAvailable.length,
      reservationValue: reservations.reduce((sum, r) => sum + Number(r.totalPrice || 0), 0),
      closedValue: closed.reduce((sum, r) => sum + Number(r.totalPrice || 0), 0)
    },
    groups: {
      reservations,
      closed,
      returnedAvailable
    }
  };
}
function inventoryMovementHourlyJob_() {
  const cfg = INVENTORY_SOURCES.EDEN_WALK;
  if (!cfg) return;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheets = inventorySheets_(ss, cfg);
  const rows = sheets.reduce((a,s)=>a.concat(inventoryReadSheetRows_(s,cfg,'EDEN_WALK')),[]);
  inventoryTrackMovements_(cfg,'EDEN_WALK',rows,ss);
  CacheService.getScriptCache().remove('DASH_V1_INV_EDEN_WALK');
}

function ensureInventoryMovementTrigger_(authToken) {
  validateAuthToken_(authToken);
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'inventoryMovementHourlyJob_');
  if (!exists) ScriptApp.newTrigger('inventoryMovementHourlyJob_').timeBased().everyHours(1).create();
  // Establish the baseline immediately. No fake movement is logged on first run.
  inventoryMovementHourlyJob_();
  return {ok:true, alreadyExisted:exists};
}
