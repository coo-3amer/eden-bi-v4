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
  const sheet = getDealSheet_();
  const existing = findExistingActiveDeal_(sheet, project, unitCode);

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
    const sheet = getDealSheet_();

    // Fail closed if the Transactions template has moved.
    assertDealWriteLayout_(sheet);

    const row = nextDealRow_(sheet);
    const seq = nextDealSequence_(sheet);
    const transactionDate = dealDate_(payload.transactionDate) || new Date();
    const paymentDate = dealDate_(payload.paymentDate);
    const birthDate = dealDate_(payload.birthDate);
    const dp = dealNumber_(payload.actualDP);
    const dpPaid = dealNumber_(payload.dpPaid);
    const actualPaid = dealNumber_(payload.actualPaid || payload.dpPaid);
    const remain = payload.remain === '' || payload.remain == null ? Math.max(0, dp - actualPaid) : dealNumber_(payload.remain);

    prepareDealRow_(sheet, row);

    // Keep the visual separator columns empty before writing any values.
    clearDealSeparatorColumns_(sheet, row);

    // Core identifiers and reservation details
    setDealCell_(sheet, row, 'A', seq.serial);
    setDealCell_(sheet, row, 'B', dealText_(payload.clientName));
    setDealCell_(sheet, row, 'D', project);
    setDealCell_(sheet, row, 'E', unitCode);
    setDealCell_(sheet, row, 'G', dealText_(payload.status));
    setDealCell_(sheet, row, 'H', dealText_(payload.additionalStatus));
    setDealCell_(sheet, row, 'I', dealNumber_(payload.reservationAmount));
    setDealCell_(sheet, row, 'J', seq.code);
    setDealCell_(sheet, row, 'K', transactionDate);
    setDealCell_(sheet, row, 'L', Utilities.formatDate(transactionDate, Session.getScriptTimeZone(), 'MMM'));
    setDealCell_(sheet, row, 'R', dealText_(payload.transactionType || 'Installment'));

    // Client information
    setDealCell_(sheet, row, 'X', dealText_(payload.idType));
    setDealCell_(sheet, row, 'Y', dealText_(payload.nationality));
    setDealCell_(sheet, row, 'Z', dealText_(payload.idNumber));
    setDealCell_(sheet, row, 'AA', dealText_(payload.idAddress));
    setDealCell_(sheet, row, 'AB', dealText_(payload.residenceAddress));
    setDealCell_(sheet, row, 'AC', dealText_(payload.mobile));
    setDealCell_(sheet, row, 'AD', dealText_(payload.email));
    setDealCell_(sheet, row, 'AE', dealText_(payload.jobTitle));
    setDealCell_(sheet, row, 'AF', birthDate);
    setDealCell_(sheet, row, 'AH', dealText_(payload.gender));
    setDealCell_(sheet, row, 'AI', dealText_(payload.clientStatue));

    // Unit / price information from inventory
    setDealCell_(sheet, row, 'AK', dealText_(payload.block));
    setDealCell_(sheet, row, 'AL', dealText_(unit.view));
    setDealCell_(sheet, row, 'AM', dealText_(unit.floor));
    setDealCell_(sheet, row, 'AN', dealText_(payload.finishingType || 'White Frame'));
    setDealCell_(sheet, row, 'AO', dealText_(unit.type || unit.unitType));
    setDealCell_(sheet, row, 'AQ', dealNumber_(unit.area));
    setDealCell_(sheet, row, 'AT', dealNumber_(unit.meterPrice));
    setDealCell_(sheet, row, 'AU', dealNumber_(unit.totalPrice));
    setDealCell_(sheet, row, 'AW', dealNumber_(unit.meterPrice));
    setDealCell_(sheet, row, 'AX', dealNumber_(unit.totalPrice));
    setDealCell_(sheet, row, 'AY', dealText_(unit.currency || payload.currency || 'USD'));
    setDealPercent_(sheet, row, 'AZ', payload.maintenancePercent);
    setDealCell_(sheet, row, 'BA', dealNumber_(payload.maintenanceAmount));
    setDealCell_(sheet, row, 'BB', dealText_(payload.discountOffer));

    // Transaction and source
    setDealCell_(sheet, row, 'BD', dealText_(payload.branch));
    setDealCell_(sheet, row, 'BE', dealText_(payload.salesName));
    setDealCell_(sheet, row, 'BF', dealText_(payload.sharedWith));
    setDealCell_(sheet, row, 'BG', dealText_(payload.salesManager));
    setDealCell_(sheet, row, 'BH', dealText_(payload.headOfSales));
    setDealCell_(sheet, row, 'BI', dealText_(payload.cco));
    setDealCell_(sheet, row, 'BJ', dealText_(payload.dealStatus || 'Solo'));
    setDealCell_(sheet, row, 'BL', dealText_(payload.mainSource || 'Direct'));
    setDealCell_(sheet, row, 'BM', dealText_(payload.sourceType));
    setDealCell_(sheet, row, 'BN', dealText_(payload.campaignName));
    setDealCell_(sheet, row, 'BO', dealText_(payload.mediaBuyer));
    setDealCell_(sheet, row, 'BP', dealText_(payload.brokerageCompany));
    setDealCell_(sheet, row, 'BQ', dealText_(payload.bcSales));

    // DP and plan
    setDealPercent_(sheet, row, 'BS', payload.dpPercent);
    setDealCell_(sheet, row, 'BT', dp);
    setDealCell_(sheet, row, 'BU', dpPaid);
    setDealCell_(sheet, row, 'BV', actualPaid);
    setDealCell_(sheet, row, 'BW', remain);
    setDealCell_(sheet, row, 'BX', paymentDate);
    setDealCell_(sheet, row, 'BZ', dealText_(payload.paymentPlanType));
    setDealCell_(sheet, row, 'CA', dealText_(payload.installments));
    setDealCell_(sheet, row, 'CC', dealText_(payload.installmentPlan));
    setDealCell_(sheet, row, 'CD', dealText_(payload.periodType));

    // Notes – final column in the approved A:DN range
    setDealCell_(sheet, row, 'DN', dealText_(payload.notes));

    SpreadsheetApp.flush();

    return {
      success: true,
      rowNumber: row,
      code: seq.code,
      message: `Deal ${seq.code} was saved successfully in row ${row}.`,
      createdBy: user.name || user.username
    };
  } finally {
    lock.releaseLock();
  }
}
