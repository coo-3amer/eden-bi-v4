const EDEN_STABILITY_DEAL_ENTRY_CONFIG = {
  spreadsheetId: '1H3Z19-Z1FJsd9rt6K6I4I9xfJZ4VFi_63gV09g7u1Ek',
  sheetName: '2026',
  headerRow: 3,
  dataStartRow: 34,
  lastColumn: 118 // DN
};

const ALL_NATIONALITIES = ["Afghan", "Albanian", "Algerian", "American", "Andorran", "Angolan", "Antiguan and Barbudan", "Argentine", "Armenian", "Australian", "Austrian", "Azerbaijani", "Bahamian", "Bahraini", "Bangladeshi", "Barbadian", "Belarusian", "Belgian", "Belizean", "Beninese", "Bhutanese", "Bolivian", "Bosnian and Herzegovinian", "Botswanan", "Brazilian", "British", "Bruneian", "Bulgarian", "Burkinabe", "Burundian", "Cabo Verdean", "Cambodian", "Cameroonian", "Canadian", "Central African", "Chadian", "Chilean", "Chinese", "Colombian", "Comorian", "Congolese", "Costa Rican", "Croatian", "Cuban", "Cypriot", "Czech", "Danish", "Djiboutian", "Dominican", "Dominican (Commonwealth)", "Dutch", "Ecuadorian", "Egyptian", "Emirati", "Equatorial Guinean", "Eritrean", "Estonian", "Eswatini", "Ethiopian", "Fijian", "Filipino", "Finnish", "French", "Gabonese", "Gambian", "Georgian", "German", "Ghanaian", "Greek", "Grenadian", "Guatemalan", "Guinean", "Guinea-Bissauan", "Guyanese", "Haitian", "Honduran", "Hungarian", "Icelandic", "Indian", "Indonesian", "Iranian", "Iraqi", "Irish", "Israeli", "Italian", "Ivorian", "Jamaican", "Japanese", "Jordanian", "Kazakhstani", "Kenyan", "Kiribati", "Kuwaiti", "Kyrgyzstani", "Lao", "Latvian", "Lebanese", "Basotho", "Liberian", "Libyan", "Liechtensteiner", "Lithuanian", "Luxembourgish", "Malagasy", "Malawian", "Malaysian", "Maldivian", "Malian", "Maltese", "Marshallese", "Mauritanian", "Mauritian", "Mexican", "Micronesian", "Moldovan", "Monegasque", "Mongolian", "Montenegrin", "Moroccan", "Mozambican", "Myanma", "Namibian", "Nauruan", "Nepalese", "New Zealander", "Nicaraguan", "Nigerien", "Nigerian", "North Korean", "North Macedonian", "Norwegian", "Omani", "Pakistani", "Palauan", "Palestinian", "Panamanian", "Papua New Guinean", "Paraguayan", "Peruvian", "Polish", "Portuguese", "Qatari", "Romanian", "Russian", "Rwandan", "Kittitian and Nevisian", "Saint Lucian", "Saint Vincentian", "Salvadoran", "Sammarinese", "Sao Tomean", "Saudi Arabian", "Senegalese", "Serbian", "Seychellois", "Sierra Leonean", "Singaporean", "Slovak", "Slovenian", "Solomon Islander", "Somali", "South African", "South Korean", "South Sudanese", "Spanish", "Sri Lankan", "Sudanese", "Surinamese", "Swedish", "Swiss", "Syrian", "Taiwanese", "Tajikistani", "Tanzanian", "Thai", "Timorese", "Togolese", "Tongan", "Trinidadian and Tobagonian", "Tunisian", "Turkish", "Turkmen", "Tuvaluan", "Ugandan", "Ukrainian", "Uruguayan", "Uzbekistani", "Vanuatuan", "Vatican", "Venezuelan", "Vietnamese", "Yemeni", "Zambian", "Zimbabwean"];

function dealColumn_(letter) {
  let n = 0;
  String(letter).toUpperCase().split('').forEach(ch => n = n * 26 + ch.charCodeAt(0) - 64);
  return n;
}

function getDealSheet_() {
  const ss = SpreadsheetApp.openById(EDEN_STABILITY_DEAL_ENTRY_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(EDEN_STABILITY_DEAL_ENTRY_CONFIG.sheetName);
  if (!sheet) throw new Error('Transactions tab "2026" was not found.');
  return sheet;
}

function dealText_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
}

function dealNumber_(v) {
  if (typeof v === 'number') return v;
  const clean = String(v || '').replace(/,/g, '').replace(/[^0-9.\-]/g, '');
  const n = Number(clean);
  return isNaN(n) ? 0 : n;
}

function dealDate_(v) {
  if (!v) return '';
  const d = v instanceof Date ? v : new Date(v);
  return isNaN(d.getTime()) ? '' : d;
}

function uniqueSortedDealValues_(values) {
  return [...new Set((values || []).map(dealText_).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function getDropDownSourceSheet_() {
  const ss = SpreadsheetApp.openById(EDEN_STABILITY_DEAL_ENTRY_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName('DROP DOWNS SOURCE');
  if (!sheet) throw new Error('DROP DOWNS SOURCE tab was not found.');
  return sheet;
}

function getDealEntryOptions(authToken) {
  validateAuthToken_(authToken);

  const source = getDropDownSourceSheet_();
  const lastRow = Math.max(source.getLastRow(), 3);
  const values = source.getRange(3, 1, lastRow - 2, 21).getDisplayValues(); // A:U

  const col = index => uniqueSortedDealValues_(values.map(r => r[index]));

  return {
    projects: col(0),              // A Projects
    branches: col(1),              // B Branches
    sales: col(2),                 // C Sales
    salesManagers: col(3),         // D Team Leader
    headOfSales: col(4),           // E Head Of Sales
    cco: col(5),                   // F CCO
    offers: col(6),                // G Offers
    statuses: col(7),              // H Status
    transactionTypes: col(8),      // I Unit Statue / Cash or Installment
    genders: col(9),               // J Gender
    clientStatuses: col(10),       // K Client statue
    mainSources: col(11),          // L Source
    sourceTypes: col(12),          // M Type
    idTypes: col(13),              // N ID
    installmentPeriods: col(14),   // O Installments Period
    campaigns: col(15),            // P Campaigns
    currencies: col(16),           // Q Currency
    brokerCompanies: col(17),      // R Broker Company
    mediaBuyers: col(18),          // S Media Buyer
    paymentPlanTypes: col(19),     // T Type Of Payment
    nationalities: uniqueSortedDealValues_(ALL_NATIONALITIES.concat(col(20))), // All nationalities + existing values
    dealStatuses: ['Solo', 'Share'],
    periodTypes: ['Month', 'Quarter', 'Semi Annual', 'Annual']
  };
}

function getDealInventoryUnits(authToken, project) {
  validateAuthToken_(authToken);
  const data = getInventoryData(authToken, project || 'KOBULETI');
  return (data.rows || []).map(r => ({
    project: r.project,
    unitCode: r.unitCode,
    status: r.status,
    company: r.company || '',
    floor: r.floor || '',
    unitType: r.unitType || '',
    type: r.type || '',
    area: r.area || 0,
    totalGross: r.totalGross || '',
    view: r.view || '',
    meterPrice: r.meterPrice || 0,
    totalPrice: r.totalPrice || 0,
    currency: r.currency || 'USD'
  }));
}


function dealNormalizeHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function dealHeaderMap_(sheet) {
  const headerRow = (typeof HEADER_ROW !== 'undefined' ? HEADER_ROW : 3);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const k = dealNormalizeHeader_(h);
    if (k && map[k] === undefined) map[k] = i + 1;
  });
  return map;
}

function dealHeaderColumn_(map, names) {
  for (const name of (names || [])) {
    const col = map[dealNormalizeHeader_(name)];
    if (col !== undefined) return col;
  }
  return 0;
}

function dealCancelledStatus_(value) {
  const s = String(value || '').trim().toLowerCase();
  return s === 'cancel' || s === 'cancelled' || s === 'canceled';
}

function assertDealTemplate_(sheet) {
  const map = dealHeaderMap_(sheet);
  const required = {
    clientName: ['Client Name', 'Final Client Name'],
    project: ['Project'],
    unitCode: ['Unit Code'],
    status: ['Status'],
    date: ['Date']
  };
  const missing = Object.keys(required).filter(k => !dealHeaderColumn_(map, required[k]));
  if (missing.length) {
    throw new Error(
      'Transactions template changed. Save was blocked to protect data. Missing headers: ' +
      missing.join(', ')
    );
  }
  return map;
}

function assertDealWriteLayout_(sheet) {
  // The current writer still uses approved fixed template columns.
  // Block saving instead of corrupting data if core columns move.
  const map = assertDealTemplate_(sheet);
  const expected = {
    'Client Name': dealColumn_('B'),
    'Project': dealColumn_('D'),
    'Unit Code': dealColumn_('E'),
    'Status': dealColumn_('G'),
    'Date': dealColumn_('K')
  };
  Object.keys(expected).forEach(name => {
    const actual = dealHeaderColumn_(map, [name, name === 'Client Name' ? 'Final Client Name' : name]);
    if (actual && actual !== expected[name]) {
      throw new Error(
        'Transactions columns were moved. Save was blocked to protect data. ' +
        name + ' is now column ' + actual + '.'
      );
    }
  });
  return map;
}

function findExistingActiveDeal_(sheet, project, unitCode) {
  const lastRow = sheet.getLastRow();
  if (lastRow < EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow) return null;

  const map = assertDealTemplate_(sheet);
  const clientCol = dealHeaderColumn_(map, ['Final Client Name', 'Client Name']);
  const projectCol = dealHeaderColumn_(map, ['Project']);
  const unitCol = dealHeaderColumn_(map, ['Unit Code']);
  const statusCol = dealHeaderColumn_(map, ['Status']);
  const additionalCol = dealHeaderColumn_(map, ['Additional Status']);
  const codeCol = dealHeaderColumn_(map, ['Res Code', 'Reservation Code', 'Code']);

  const lastCol = sheet.getLastColumn();
  const count = lastRow - EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow + 1;
  const values = sheet.getRange(EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow, 1, count, lastCol).getDisplayValues();
  const projectKey = dealText_(project).toLowerCase();
  const unitKey = dealText_(unitCode).toLowerCase();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowProject = dealText_(row[projectCol - 1]).toLowerCase();
    const rowUnit = dealText_(row[unitCol - 1]).toLowerCase();
    const status = dealText_(row[statusCol - 1]);
    const additional = additionalCol ? dealText_(row[additionalCol - 1]).toLowerCase() : '';
    const code = codeCol ? dealText_(row[codeCol - 1]) : '';

    if (rowProject !== projectKey || rowUnit !== unitKey) continue;
    if (additional.includes('renovation') || code.toLowerCase().includes('reno')) continue;
    if (dealCancelledStatus_(status)) continue;

    return {
      rowNumber: EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow + i,
      clientName: clientCol ? dealText_(row[clientCol - 1]) : '',
      project: rowProject,
      unitCode: rowUnit,
      status,
      code
    };
  }
  return null;
}

function checkDealUnitAvailability(authToken, project, unitCode) {
  validateAuthToken_(authToken);
  project = dealText_(project || 'KOBULETI');
  unitCode = dealText_(unitCode);
  if (!unitCode) return { allowed: false, message: 'Select a unit first.' };

  const inventoryCheck = validateInventoryUnitForDeal(authToken, project, unitCode);
  const mainSheet = getMainDealsSheet_();
  const existing = findExistingActiveMainDeal_(mainSheet, mainDealsColumns_(mainSheet), project, unitCode);

  if (existing) {
    return {
      allowed: false,
      reason: 'EXISTING_ACTIVE_DEAL',
      message: 'This unit already has an active transaction.',
      existing: existing,
      unit: inventoryCheck.unit || null,
      inventoryStatus: inventoryCheck.status || ''
    };
  }

  if (!inventoryCheck.allowed) {
    return {
      allowed: false,
      reason: inventoryCheck.reason,
      message: inventoryCheck.message,
      unit: inventoryCheck.unit || null,
      inventoryStatus: inventoryCheck.status || ''
    };
  }

  return {
    allowed: true,
    message: 'Unit is available and can be added.',
    unit: inventoryCheck.unit,
    inventoryStatus: inventoryCheck.status
  };
}

function nextDealRow_(sheet) {
  const start = EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow;
  const last = Math.max(sheet.getLastRow(), start);
  const count = last - start + 1;
  const keys = sheet.getRange(start, dealColumn_('B'), count, 3).getDisplayValues(); // B:D
  for (let i = 0; i < keys.length; i++) {
    if (!dealText_(keys[i][0]) && !dealText_(keys[i][2])) return start + i;
  }
  return last + 1;
}

function nextDealSequence_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow) return { serial: 1, code: 'ED-REV-01' };
  const count = lastRow - EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow + 1;
  const serials = sheet.getRange(EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow, dealColumn_('A'), count, 1).getDisplayValues().flat();
  const codes = sheet.getRange(EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow, dealColumn_('J'), count, 1).getDisplayValues().flat();

  const maxSerial = serials.reduce((m, v) => Math.max(m, parseInt(String(v).replace(/\D/g, ''), 10) || 0), 0);
  const maxCode = codes.reduce((m, v) => {
    const match = String(v || '').match(/ED-REV-(\d+)/i);
    return Math.max(m, match ? Number(match[1]) : 0);
  }, 0);

  return {
    serial: maxSerial + 1,
    code: 'ED-REV-' + String(maxCode + 1).padStart(2, '0')
  };
}

function prepareDealRow_(sheet, rowNumber) {
  if (rowNumber > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rowNumber - sheet.getMaxRows());
  }

  const sourceRow = Math.max(EDEN_STABILITY_DEAL_ENTRY_CONFIG.dataStartRow, rowNumber - 1);
  if (sourceRow !== rowNumber && sourceRow <= sheet.getLastRow()) {
    const src = sheet.getRange(sourceRow, 1, 1, EDEN_STABILITY_DEAL_ENTRY_CONFIG.lastColumn);
    const dst = sheet.getRange(rowNumber, 1, 1, EDEN_STABILITY_DEAL_ENTRY_CONFIG.lastColumn);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

    const formulas = src.getFormulasR1C1()[0];
    formulas.forEach((formula, i) => {
      if (formula) sheet.getRange(rowNumber, i + 1).setFormulaR1C1(formula);
    });
  }
}

function setDealCell_(sheet, row, letter, value) {
  if (value === undefined || value === null || value === '') return;
  sheet.getRange(row, dealColumn_(letter)).setValue(value);
}

function setDealPercent_(sheet, row, letter, value) {
  if (value === undefined || value === null || value === '') return;
  const raw = dealNumber_(value);
  const normalized = Math.abs(raw) > 1 ? raw / 100 : raw;
  sheet.getRange(row, dealColumn_(letter))
    .setValue(normalized)
    .setNumberFormat('0.00%');
}

function clearDealSeparatorColumns_(sheet, row) {
  // These columns are intentionally blank in the Transactions template
  // and visually separate the sheet sections.
  const separators = [
    'C','F','O','W','AJ','AP','BC','BK','BR','BY',
    'CB','CE','CF','CR','CW','DB','DG'
  ];
  separators.forEach(letter => sheet.getRange(row, dealColumn_(letter)).clearContent());
}

function saveNewDeal(authToken, payload) {
  const user = validateAuthToken_(authToken);

  // SECURITY: UI hiding is not authorization.
  if (typeof isEgyptViewerUser_ === 'function' && isEgyptViewerUser_(user)) {
    throw new Error('This account is read only.');
  }

  payload = payload || {};

  const required = ['clientName', 'project', 'unitCode', 'status', 'mobile', 'salesName', 'branch'];
  const missing = required.filter(k => !dealText_(payload[k]));
  if (missing.length) throw new Error('Complete the required fields: ' + missing.join(', '));

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const project = dealText_(payload.project).toUpperCase();
    const unitCode = dealText_(payload.unitCode);
    const check = checkDealUnitAvailability(authToken, project, unitCode);
    if (!check.allowed) {
      const details = check.existing ? ` Existing client: ${check.existing.clientName}. Status: ${check.existing.status}.` : '';
      throw new Error(check.message + details);
    }

    const unit = check.unit || {};
    const saved = saveDealToMainSheet_(user, Object.assign({}, payload, { project: project, unitCode: unitCode }), unit);
    clearDashboardCache_();
    const ref = saved.dealNum ? 'Deal #' + saved.dealNum : 'The deal';
    return {
      success: true,
      rowNumber: saved.row,
      code: String(saved.dealNum || ''),
      written: saved.written,
      skipped: saved.skipped,
      message: `${ref} was saved in the Deals sheet, row ${saved.row}.`,
      createdBy: user.name || user.username
    };
  } finally {
    lock.releaseLock();
  }
}


/* =====================================================================
 * MAIN DEALS SHEET WRITER
 * New deals from the system go to the main Deals sheet (the one the
 * dashboard reads), in the same shape as the rows already there:
 *  - the new row is placed right after the last deal,
 *  - formats, dropdowns and formulas are copied from the row above,
 *  - every value is written by its column header (never by letter),
 *  - cells that hold a formula are left to calculate on their own,
 *  - the Project cell gets a note "EDEN BI • …" so the dashboard knows
 *    the row came from the system.
 * The empty rows under the data carry a warning-only protection, so
 * anyone typing a new deal straight into the sheet is told to use the
 * system (run setupDealsSheetGuard once from the editor).
 * ===================================================================== */
const DEALS_GUARD_DESC_ = 'EDEN BI — add new deals from the dashboard (Add New Deal)';
const DEALS_SYSTEM_NOTE_ = 'EDEN BI';

function getMainDealsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByGid_(ss, DEALS_GID) || ss.getSheetByName('Deals');
  if (!sheet) throw new Error('The Deals sheet was not found.');
  return sheet;
}

/* header (normalized) -> column number; grouped sub-headers by "group|sub". */
function mainDealsColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const k = normalizeHeader_(h);
    if (k && map[k] === undefined) map[k] = i + 1;
  });
  const group = (name, subs) => {
    const g = groupedHeaderIndexes_(sheet, Math.max(1, HEADER_ROW - 1), HEADER_ROW, name, subs);
    Object.keys(g).forEach(s => { map[normalizeHeader_(name) + '|' + normalizeHeader_(s)] = g[s] + 1; });
  };
  try { group('Client Info', ['Contact']); } catch (e) {}
  try { group('Unit Info', ['Unit Type']); } catch (e) {}
  try { group('Source Details', ['Main Source', 'Source Type', 'Campaign Name', 'Media Buyer', 'Brokerage Company', 'BC Sales', 'BC Manger']); } catch (e) {}
  return { map: map, lastCol: lastCol };
}

function mainDealsCol_(cols, names) {
  for (const n of names) {
    if (String(n).charAt(0) === '@') {           // fixed column confirmed by the sheet owner
      const c = dealColumn_(String(n).slice(1));
      if (c && c <= cols.lastCol) return c;
      continue;
    }
    const c = cols.map[normalizeHeader_(n)];
    if (c) return c;
  }
  return 0;
}

/* Last row that holds a deal (Project, Client Name or Unit Code filled). */
function mainDealsLastRow_(sheet, cols) {
  const last = sheet.getLastRow();
  if (last < DATA_START_ROW) return DATA_START_ROW - 1;
  const keyCols = [
    mainDealsCol_(cols, ['Project']),
    mainDealsCol_(cols, ['Client Name', 'Final Client Name']),
    mainDealsCol_(cols, ['Unit Code'])
  ].filter(Boolean);
  const n = last - DATA_START_ROW + 1;
  let found = DATA_START_ROW - 1;
  keyCols.forEach(c => {
    const v = sheet.getRange(DATA_START_ROW, c, n, 1).getDisplayValues();
    for (let i = v.length - 1; i >= 0; i--) {
      if (String(v[i][0]).trim()) { found = Math.max(found, DATA_START_ROW + i); break; }
    }
  });
  return found;
}

function findExistingActiveMainDeal_(sheet, cols, project, unitCode) {
  const last = mainDealsLastRow_(sheet, cols);
  if (last < DATA_START_ROW) return null;
  const pc = mainDealsCol_(cols, ['Project']), uc = mainDealsCol_(cols, ['Unit Code']);
  const sc = mainDealsCol_(cols, ['Status']), ac = mainDealsCol_(cols, ['Additional Status']);
  const cc = mainDealsCol_(cols, ['Final Client Name', 'Client Name']);
  if (!pc || !uc) return null;
  const n = last - DATA_START_ROW + 1;
  const read = c => c ? sheet.getRange(DATA_START_ROW, c, n, 1).getDisplayValues().map(r => r[0]) : new Array(n).fill('');
  const P = read(pc), U = read(uc), S = read(sc), A = read(ac), C = read(cc);
  const pk = dealText_(project).toLowerCase(), uk = dealText_(unitCode).toLowerCase();
  for (let i = 0; i < n; i++) {
    if (dealText_(P[i]).toLowerCase() !== pk || dealText_(U[i]).toLowerCase() !== uk) continue;
    if (dealText_(A[i]).toLowerCase().includes('renovation')) continue;
    if (dealCancelledStatus_(S[i])) continue;
    return { rowNumber: DATA_START_ROW + i, clientName: dealText_(C[i]), project: pk, unitCode: uk, status: dealText_(S[i]), code: '' };
  }
  return null;
}

/* Copies format, dropdowns and formulas of the row above into the new row. */
function prepareMainDealRow_(sheet, cols, row) {
  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows() + 50);
  // Never overwrite something already typed in the target row (e.g. a totals row).
  const target = sheet.getRange(row, 1, 1, cols.lastCol);
  const tf = target.getFormulas()[0];
  const hasValue = target.getDisplayValues()[0].some((v, i) => String(v).trim() && !tf[i]);
  if (hasValue) sheet.insertRowBefore(row);
  const src = sheet.getRange(row - 1, 1, 1, cols.lastCol);
  if (row - 1 >= DATA_START_ROW) {
    const dst = sheet.getRange(row, 1, 1, cols.lastCol);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    const f = src.getFormulasR1C1()[0];
    f.forEach((formula, i) => { if (formula) sheet.getRange(row, i + 1).setFormulaR1C1(formula); });
    try { sheet.setRowHeight(row, sheet.getRowHeight(row - 1)); } catch (e) {}
  }
  return src.getDisplayValues()[0];
}

/* Moves the warning-only protection so it always covers the empty rows. */
function updateDealsGuard_(sheet, lastDealRow) {
  const start = lastDealRow + 1;
  if (start > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
  const range = sheet.getRange(start, 1, sheet.getMaxRows() - start + 1, sheet.getMaxColumns());
  let p = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(x => x.getDescription() === DEALS_GUARD_DESC_)[0];
  if (!p) p = range.protect().setDescription(DEALS_GUARD_DESC_);
  else p.setRange(range);
  p.setWarningOnly(true);
  return start;
}

/* Run once from the Apps Script editor. */
function setupDealsSheetGuard() {
  const sheet = getMainDealsSheet_();
  const cols = mainDealsColumns_(sheet);
  const start = updateDealsGuard_(sheet, mainDealsLastRow_(sheet, cols));
  Logger.log('Deals sheet: rows from ' + start + ' down now show a warning. New deals should be added from the dashboard.');
}

/* Run from the editor to see which form fields match which Deals columns. */
function checkDealsSheetMapping() {
  const sheet = getMainDealsSheet_();
  const cols = mainDealsColumns_(sheet);
  const head = sheet.getRange(HEADER_ROW, 1, 1, cols.lastCol).getDisplayValues()[0];
  const lines = MAIN_DEAL_FIELDS_.map(f => {
    const c = mainDealsCol_(cols, f.h);
    const name = f.h.filter(x => x.indexOf('|') < 0 && x.charAt(0) !== '@')[0] || f.h[0];
    return (c ? 'OK   ' + columnLetter_(c) + '  ' : 'MISS     ') + name + (c ? '   ← sheet header: "' + String(head[c - 1]).replace(/\s+/g, ' ').trim() + '"' : '');
  });
  Logger.log('Last deal row: ' + mainDealsLastRow_(sheet, cols) + '\n' + lines.join('\n'));
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/* form value -> Deals column header(s). First header found is used. */
const MAIN_DEAL_FIELDS_ = [
  { k: 'dealNum',          h: ['Deal Num', 'Deal Number', 'Deal No'] },
  { k: 'clientName',       h: ['Client Name'] },
  { k: 'project',          h: ['Project'] },
  { k: 'unitCode',         h: ['Unit Code'] },
  { k: 'status',           h: ['Status'] },
  { k: 'additionalStatus', h: ['Additional Status'] },
  { k: 'reservationAmount',h: ['Reservation Amount', 'Res Amount', '@J'] },
  { k: 'date',             h: ['Date'] },
  { k: 'month',            h: ['Month'] },
  { k: 'year',             h: ['Year'] },
  { k: 'mobile',           h: ['client info|contact', 'Contact', 'Mobile', 'Mobile Number', 'Phone'] },
  { k: 'clientType',       h: ['Client Type'] },
  { k: 'email',            h: ['Email', 'E-mail', 'Mail'] },
  { k: 'nationality',      h: ['Nationality', 'Nationalty'] },
  { k: 'idType',           h: ['ID Type', 'ID'] },
  { k: 'idNumber',         h: ['ID Number', 'ID No', 'National ID', 'Passport Number'] },
  { k: 'idAddress',        h: ['ID Address', 'Address'] },
  { k: 'residenceAddress', h: ['Residence Address', 'Current Address'] },
  { k: 'jobTitle',         h: ['Job Title', 'Job'] },
  { k: 'birthDate',        h: ['Birth Date', 'Date Of Birth', 'DOB'] },
  { k: 'gender',           h: ['Gender'] },
  { k: 'clientStatue',     h: ['Client Statue', 'Client Status'] },
  { k: 'block',            h: ['Block'] },
  { k: 'view',             h: ['View'] },
  { k: 'floor',            h: ['Floor'] },
  { k: 'finishing',        h: ['Finishing Type', 'Finishing'] },
  { k: 'unitType',         h: ['unit info|unit type', 'Unit Type'] },
  { k: 'area',             h: ['Area'] },
  { k: 'primaryMeter',     h: ['Primary Meter Price'] },
  { k: 'primaryTotal',     h: ['Primary Total Price'] },
  { k: 'meterAfter',       h: ['Meter Price After Discount'] },
  { k: 'finalPrice',       h: ['Final Price'] },
  { k: 'currency',         h: ['Currency'] },
  { k: 'maintenancePercent', h: ['Maintenance %', 'Maintenance Percent'] },
  { k: 'maintenanceAmount',h: ['Maintenance Amount'] },
  { k: 'discountOffer',    h: ['Discount Offer', 'Offer'] },
  { k: 'branch',           h: ['Branch'] },
  { k: 'salesName',        h: ['Sales Name'] },
  { k: 'sharedWith',       h: ['Shared With'] },
  { k: 'salesManager',     h: ['Sales Manager'] },
  { k: 'headOfSales',      h: ['Head Of Sales'] },
  { k: 'cco',              h: ['CCO'] },
  { k: 'dealStatus',       h: ['Deal Status'] },
  { k: 'mainSource',       h: ['source details|main source', 'Main Source'] },
  { k: 'sourceType',       h: ['source details|source type', 'Source Type'] },
  { k: 'campaignName',     h: ['source details|campaign name', 'Campaign Name'] },
  { k: 'mediaBuyer',       h: ['source details|media buyer', 'Media Buyer'] },
  { k: 'brokerageCompany', h: ['source details|brokerage company', 'Brokerage Company'] },
  { k: 'bcSales',          h: ['source details|bc sales', 'BC Sales'] },
  { k: 'dpPercent',        h: ['DP %', 'DP%'] },
  { k: 'actualDP',         h: ['Actual DP', 'Acctual DP'] },
  { k: 'dpPaid',           h: ['DP Paid'] },
  { k: 'actualPaid',       h: ['Actual Paid', 'Acctual Paid'] },
  { k: 'remain',           h: ['Remain'] },
  { k: 'paymentDate',      h: ['Payment Date', '@CK'] },
  { k: 'paymentPlanType',  h: ['Type Of Payment', 'Payment Plan', '@CN'] },
  { k: 'installments',     h: ['Installments Period', 'Installments'] },
  { k: 'installmentPlan',  h: ['Installment Plan'] },
  { k: 'periodType',       h: ['Period Type', 'Period'] },
  { k: 'notes',            h: ['Notes', 'Note', 'Comments', '@EE'] }
];

function saveDealToMainSheet_(user, payload, unit) {
  const sheet = getMainDealsSheet_();
  const cols = mainDealsColumns_(sheet);
  const need = { 'Project': ['Project'], 'Unit Code': ['Unit Code'], 'Status': ['Status'], 'Client Name': ['Client Name', 'Final Client Name'], 'Date': ['Date'] };
  const missing = Object.keys(need).filter(n => !mainDealsCol_(cols, need[n]));
  if (missing.length) throw new Error('Deals sheet columns not found (' + missing.join(', ') + '). Nothing was saved.');

  const last = mainDealsLastRow_(sheet, cols);
  let row = last + 1;
  const above = prepareMainDealRow_(sheet, cols, row);
  // insertRowBefore may have pushed a typed row down; our row number stays the same.

  const tz = Session.getScriptTimeZone();
  const date = dealDate_(payload.transactionDate) || new Date();
  const monthCol = mainDealsCol_(cols, ['Month']);
  const monthAbove = monthCol ? String(above[monthCol - 1] || '').trim() : '';
  const month = /^\d+$/.test(monthAbove) ? date.getMonth() + 1
    : (monthAbove.length > 3 ? Utilities.formatDate(date, tz, 'MMMM') : Utilities.formatDate(date, tz, 'MMM'));

  // Next deal number
  let dealNum = '';
  const numCol = mainDealsCol_(cols, ['Deal Num', 'Deal Number', 'Deal No']);
  if (numCol && last >= DATA_START_ROW) {
    const nums = sheet.getRange(DATA_START_ROW, numCol, last - DATA_START_ROW + 1, 1).getDisplayValues();
    dealNum = nums.reduce((m, r) => Math.max(m, parseInt(String(r[0]).replace(/\D/g, ''), 10) || 0), 0) + 1;
  }

  const dp = dealNumber_(payload.actualDP);
  const actualPaid = dealNumber_(payload.actualPaid || payload.dpPaid);
  const pct = v => { if (v === '' || v == null) return ''; const n = dealNumber_(v); return Math.abs(n) > 1 ? n / 100 : n; };
  const price = dealNumber_(unit.totalPrice), meter = dealNumber_(unit.meterPrice);
  const v = {
    dealNum: dealNum,
    clientName: dealText_(payload.clientName),
    finalClientName: dealText_(payload.clientName),
    project: dealText_(payload.project).toUpperCase(),
    unitCode: dealText_(payload.unitCode),
    status: dealText_(payload.status),
    additionalStatus: dealText_(payload.additionalStatus),
    reservationAmount: dealNumber_(payload.reservationAmount) || '',
    date: date,
    month: month,
    year: date.getFullYear(),
    transactionType: dealText_(payload.transactionType),
    mobile: dealText_(payload.mobile),
    clientType: (typeof classifyClientTypeFromMobileHeader_ === 'function') ? classifyClientTypeFromMobileHeader_(payload.mobile) : '',
    email: dealText_(payload.email),
    nationality: dealText_(payload.nationality),
    idType: dealText_(payload.idType),
    idNumber: dealText_(payload.idNumber),
    idAddress: dealText_(payload.idAddress),
    residenceAddress: dealText_(payload.residenceAddress),
    jobTitle: dealText_(payload.jobTitle),
    birthDate: dealDate_(payload.birthDate),
    gender: dealText_(payload.gender),
    clientStatue: dealText_(payload.clientStatue),
    block: dealText_(payload.block),
    view: dealText_(unit.view),
    floor: dealText_(unit.floor),
    finishing: dealText_(payload.finishingType),
    unitType: dealText_(unit.type || unit.unitType),
    area: dealNumber_(unit.area) || '',
    primaryMeter: meter || '',
    primaryTotal: price || '',
    meterAfter: meter || '',
    finalPrice: price || '',
    currency: dealText_(unit.currency || payload.currency),
    maintenancePercent: pct(payload.maintenancePercent),
    maintenanceAmount: dealNumber_(payload.maintenanceAmount) || '',
    discountOffer: dealText_(payload.discountOffer),
    branch: dealText_(payload.branch),
    salesName: dealText_(payload.salesName),
    sharedWith: dealText_(payload.sharedWith),
    salesManager: dealText_(payload.salesManager),
    headOfSales: dealText_(payload.headOfSales),
    cco: dealText_(payload.cco),
    dealStatus: dealText_(payload.dealStatus || 'Solo'),
    mainSource: dealText_(payload.mainSource),
    sourceType: dealText_(payload.sourceType),
    campaignName: dealText_(payload.campaignName),
    mediaBuyer: dealText_(payload.mediaBuyer),
    brokerageCompany: dealText_(payload.brokerageCompany),
    bcSales: dealText_(payload.bcSales),
    dpPercent: pct(payload.dpPercent),
    actualDP: dp || '',
    dpPaid: dealNumber_(payload.dpPaid) || '',
    actualPaid: actualPaid || '',
    remain: payload.remain === '' || payload.remain == null ? (dp ? Math.max(0, dp - actualPaid) : '') : dealNumber_(payload.remain),
    paymentDate: dealDate_(payload.paymentDate),
    paymentPlanType: dealText_(payload.paymentPlanType),
    installments: dealText_(payload.installments),
    installmentPlan: dealText_(payload.installmentPlan),
    periodType: dealText_(payload.periodType),
    notes: dealText_(payload.notes)
  };

  const formulas = sheet.getRange(row, 1, 1, cols.lastCol).getFormulas()[0];
  const used = {}, written = [], skipped = [];
  MAIN_DEAL_FIELDS_.forEach(f => {
    const value = v[f.k];
    if (value === '' || value == null) return;
    const c = mainDealsCol_(cols, f.h);
    const label = f.h.filter(x => x.indexOf('|') < 0 && x.charAt(0) !== '@')[0] || f.h[0];
    if (!c) { skipped.push(label); return; }
    if (used[c] || formulas[c - 1]) return;   // one value per column; formulas stay
    used[c] = true;
    sheet.getRange(row, c).setValue(value);
    written.push(label);
  });

  const pc = mainDealsCol_(cols, ['Project']);
  sheet.getRange(row, pc).setNote(DEALS_SYSTEM_NOTE_ + ' • added by ' + (user.name || user.username) + ' • ' +
    Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'));

  SpreadsheetApp.flush();
  try { updateDealsGuard_(sheet, row); } catch (e) {}
  return { row: row, dealNum: dealNum, written: written, skipped: skipped };
}
