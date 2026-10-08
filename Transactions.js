/**
 * EDEN BI Transactions mapping policy:
 * Business fields are resolved ONLY by column header names.
 * No fixed spreadsheet column letters or numeric row indexes are used.
 */
function normalizeClientType_(value) {
  const text = String(value || '').trim().toLowerCase();

  if (text === 'internal') return 'Internal';
  if (text === 'overseas') return 'Overseas';

  return String(value || '').trim();
}


function normalizePhoneForClientType_(value) {
  return String(value || '')
    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d))
    .replace(/\D/g, '');
}

function classifyClientTypeFromMobileHeader_(mobile) {
  const values = String(mobile || '')
    .split(/[\n\r,;|/]+/)
    .map(normalizePhoneForClientType_)
    .filter(Boolean);

  if (!values.length) return '';

  const isEgyptian = phone => {
    let p = String(phone || '');
    if (p.indexOf('0020') === 0) p = p.slice(2);
    return /^01[0125]\d{8}$/.test(p) || /^201[0125]\d{8}$/.test(p);
  };

  return values.every(isEgyptian) ? 'Internal' : 'Overseas';
}


/**
 * Resolves duplicate subheaders inside a top-level merged header group.
 * Example: Internal comm -> Sales / Managers / Staff / Statue.
 * This deliberately bypasses buildIndex_ so another "Sales" column cannot collide.
 */
function groupedHeaderIndexes_(sheet, groupRow, subHeaderRow, groupName, subHeaders) {
  const lastCol = sheet.getLastColumn();
  const groups = sheet.getRange(groupRow, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || '').trim());
  const subs = sheet.getRange(subHeaderRow, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const norm = v => String(v || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const wantedGroup = norm(groupName);
  let start = -1;

  for (let i = 0; i < groups.length; i++) {
    if (norm(groups[i]) === wantedGroup) { start = i; break; }
  }
  if (start < 0) return {};

  let end = groups.length;
  for (let i = start + 1; i < groups.length; i++) {
    if (norm(groups[i])) { end = i; break; }
  }

  const result = {};
  (subHeaders || []).forEach(name => {
    const key = norm(name);
    for (let i = start; i < end; i++) {
      if (norm(subs[i]) === key) { result[name] = i; break; }
    }
  });
  return result;
}

function cancelledStatus_(value) {
  const s = String(value || '').trim().toLowerCase();
  return s === 'cancel' || s === 'cancelled' || s === 'canceled';
}

/**
 * Dashboard data, cached for a few minutes so most visits skip re-reading
 * the sheets. opts.fresh = true bypasses the cache (Reload button).
 * The cache is cleared when a deal is saved from the dashboard.
 */
const DASH_CACHE_SECONDS_ = 21600;      // kept up to 6 h (CacheService maximum)…
const DASH_CACHE_MAX_AGE_MS_ = 60 * 60 * 1000; // …rebuilt on a visit only if older than 1 h (LiveSync + warm-up keep it fresh)
const DASH_CACHE_SLICE_ = 90000;

function getDashboardData(authToken, opts) {
  const authUser = validateAuthToken_(authToken);
  const egyptViewer = isEgyptViewerUser_(authUser);
  const scope = egyptViewer ? 'EGV' : 'ALL';
  const fresh = !!(opts && opts.fresh);
  let data = fresh ? null : dashCacheRead_(scope);
  if (!data) {
    data = buildDashboardData_(authUser);
    dashCacheWrite_(scope, data);
  }
  data.dataVersion = dashCacheVersion_(scope);   // lets the page notice newer data (LiveSync)
  if (isSalesScopedUser_(authUser)) return scopeDashboardForSales_(data, authUser);
  data.access = egyptViewerAccessProfile_(authUser);
  return data;
}

/* Fingerprint of a cached block's content ('' if missing). */
function dashCacheVersion_(scope) {
  try { const m = CacheService.getScriptCache().get('DASH_V1_' + scope); return m ? String(JSON.parse(m).h || JSON.parse(m).at || '') : ''; }
  catch (e) { return ''; }
}

/* When a cached block was built (ms), 0 if missing. */
function dashCacheAt_(scope) {
  try { const m = CacheService.getScriptCache().get('DASH_V1_' + scope); return m ? (JSON.parse(m).at || 0) : 0; }
  catch (e) { return 0; }
}

function dashCacheRead_(scope, maxAgeMs) {
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.get('DASH_V1_' + scope);
    if (!meta) return null;
    const m = JSON.parse(meta);
    if (!m.at || Date.now() - m.at > (maxAgeMs || DASH_CACHE_MAX_AGE_MS_)) return null;
    const keys = [];
    for (let i = 0; i < m.parts; i++) keys.push('DASH_V1_' + scope + '_' + m.id + '_' + i);
    const got = cache.getAll(keys);
    if (keys.some(k => got[k] == null)) return null;
    const b64 = keys.map(k => got[k]).join('');
    const text = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString('UTF-8');
    return JSON.parse(text);
  } catch (err) {
    console.warn('Dashboard cache read: ' + err);
    return null;
  }
}

function dashCacheWrite_(scope, data) {
  try {
    const copy = Object.assign({}, data);
    delete copy.access;
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(copy), 'application/json')).getBytes());
    const id = Utilities.getUuid().slice(0, 8);
    const batch = {};
    let parts = 0;
    for (let i = 0; i < b64.length; i += DASH_CACHE_SLICE_) batch['DASH_V1_' + scope + '_' + id + '_' + (parts++)] = b64.slice(i, i + DASH_CACHE_SLICE_);
    const cache = CacheService.getScriptCache();
    cache.putAll(batch, DASH_CACHE_SECONDS_ + 60);
    // h = fingerprint of the content: unchanged data keeps the same h, so open
    // dashboards only refresh when something really changed.
    const h = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([copy.rows, copy.brokers, copy.options]))).slice(0, 16);
    cache.put('DASH_V1_' + scope, JSON.stringify({ id: id, parts: parts, at: Date.now(), h: h }), DASH_CACHE_SECONDS_);
  } catch (err) {
    console.warn('Dashboard cache write: ' + err);
  }
}

/**
 * Runs every 5 minutes (time-driven trigger) so the dashboard data is always
 * ready when someone opens the dashboard: nobody waits for the sheets.
 */
function warmDashboardCache() {
  // Never run two warm-ups at the same time (a slow run used to overlap the next one).
  // A cache flag, not a script lock, so saving a deal never waits for the warm-up.
  const cache = CacheService.getScriptCache();
  if (cache.get('WARM_RUNNING')) return;
  cache.put('WARM_RUNNING', '1', 330);
  const t0 = Date.now(), elapsed = () => Date.now() - t0;
  try {
    // 1) What every visitor needs first.
    dashCacheWrite_('ALL', buildDashboardData_({ role: 'Admin' }));
    if (dashCacheAge_('EGV') > 12 * 60 * 1000) {
      try { dashCacheWrite_('EGV', buildDashboardData_({ role: 'Egypt Viewer' })); } catch (err) { console.warn('Egypt viewer warm-up: ' + err); }
    }
    // 2) Inventories (Add New Deal + Inventory pages), only when getting old, while time allows.
    const order = ['EDEN_WALK', 'CITY_CENTER_GLDANI'].concat(Object.keys(INVENTORY_SOURCES).filter(k => k !== 'EDEN_WALK' && k !== 'CITY_CENTER_GLDANI'));
    order.forEach(k => {
      if (elapsed() > 240000 || dashCacheAge_('INV_' + k) < 30 * 60 * 1000) return;
      try { inventoryDataCached_(k, true, false); } catch (e) { console.warn('Inventory warm ' + k + ': ' + e); }
    });
    // 3) Add New Deal dropdowns (hourly).
    if (elapsed() < 150000) { try { warmDealEntryOptions_(); } catch (e) { console.warn('Deal options warm: ' + e); } }
  } finally {
    cache.remove('WARM_RUNNING');
    console.log('Warm-up took ' + elapsed() + ' ms');
  }
}

/* Age in ms of a cached block (Infinity when missing). */
function dashCacheAge_(scope) {
  try {
    const meta = CacheService.getScriptCache().get('DASH_V1_' + scope);
    return meta ? Date.now() - (JSON.parse(meta).at || 0) : Infinity;
  } catch (e) { return Infinity; }
}

/**
 * Run ONCE from the Apps Script editor (select it, then Run) to start the
 * 5-minute warm-up. Running it again does not create duplicates.
 */
function setupDashboardWarmup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'warmDashboardCache')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('warmDashboardCache').timeBased().everyMinutes(5).create();
  warmDashboardCache();
  return 'Dashboard warm-up is on (every 5 minutes).';
}

function clearDashboardCache_() {
  try { CacheService.getScriptCache().removeAll(['DASH_V1_ALL', 'DASH_V1_EGV']); } catch (err) {}
}

function buildDashboardData_(authUser) {
  const egyptViewer = isEgyptViewerUser_(authUser);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet =
    getSheetByGid_(ss, DEALS_GID) ||
    ss.getSheetByName('Deals') ||
    ss.getSheets()[0];

  if (!sheet) throw new Error('لم يتم العثور على شيت الداتا');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < DATA_START_ROW) {
    return {
      sheetName: sheet.getName(),
      rows: [],
      options: {},
      brokers: getBrokerDashboardData_(),
      access: egyptViewerAccessProfile_(authUser)
    };
  }

  const headers = sheet
    .getRange(HEADER_ROW, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(h => String(h || '').trim());

  const displayRows = sheet
    .getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol)
    .getDisplayValues();

  const idx = buildIndex_(headers);

  // CLIENT DOCUMENTS: the client name cells link to each client's papers.
  const clientDocLinks = clientDocumentLinks_(sheet, idx, displayRows.length);

  // Rows added from the dashboard carry an "EDEN BI" note on the Project cell.
  const systemAdded = dealsSystemNotes_(sheet, headers, displayRows.length);

  // STABILITY V1: resolve Internal comm subheaders inside their own group.
  const internalCommIdx = groupedHeaderIndexes_(
    sheet,
    Math.max(1, HEADER_ROW - 1),
    HEADER_ROW,
    'Internal comm',
    ['Sales', 'Managers', 'Staff', 'Statue']
  );

  // CONTRACT DETAILS: the contract Date sits under its own group (not the reservation Date).
  const contractDetailsIdx = groupedHeaderIndexes_(
    sheet,
    Math.max(1, HEADER_ROW - 1),
    HEADER_ROW,
    'Contract Details',
    ['Date', 'Client Sign Date', 'Company Sign Date']
  );

  // DOWN PAYMENT / PAYMENT PLAN: used to prefill a unit's payment schedule.
  const dpDetailsIdx = groupedHeaderIndexes_(sheet, Math.max(1, HEADER_ROW - 1), HEADER_ROW, 'Down Payment Details', ['Date']);
  const planIdx = groupedHeaderIndexes_(sheet, Math.max(1, HEADER_ROW - 1), HEADER_ROW, 'Payment Plan Details',
    ['Type Of Payment Plan', 'Installments Period', 'Installment Plan', 'Period Type']);

  // DEAL CONFIRMATION: contract-stage checkboxes (client already signed).
  const confIdx = groupedHeaderIndexes_(sheet, Math.max(1, HEADER_ROW - 1), HEADER_ROW, 'Deal Confirmation',
    ['Part of DP', 'Total DP', 'Client Sigend', 'Company Signed', 'Delivery Signed', 'Delivered']);

  // UNIT INFO: resolve the Unit Type subheader inside its merged group.
  // Confirmed sheet structure: Unit Info -> Unit Type (AY).
  const unitInfoIdx = groupedHeaderIndexes_(
    sheet,
    Math.max(1, HEADER_ROW - 1),
    HEADER_ROW,
    'Unit Info',
    ['Unit Type']
  );

  // CLIENT INFO: resolve Contact by grouped header names only.
  const clientInfoIdx = groupedHeaderIndexes_(
    sheet,
    Math.max(1, HEADER_ROW - 1),
    HEADER_ROW,
    'Client Info',
    ['Contact']
  );

  // SOURCE DETAILS: resolve all source fields by grouped header names only.
  // No column letters or numeric positions are used.
  const sourceDetailsIdx = groupedHeaderIndexes_(
    sheet,
    Math.max(1, HEADER_ROW - 1),
    HEADER_ROW,
    'Source Details',
    [
      'Main Source',
      'Source Type',
      'Campaign Name',
      'Media Buyer',
      'Brokerage Company',
      'BC Sales',
      'BC Manger'
    ]
  );

  const hasHeader = name =>
    idx[name] !== undefined ||
    idx[normalizeHeader_(name)] !== undefined;

  const availableFilters = {
    project: hasHeader('Project'),
    unitCategory:
      hasHeader('Unit Type') ||
      hasHeader('Unit Category') ||
      hasHeader('UnitType') ||
      hasHeader('Unit_Type') ||
      hasHeader('Type'),
    sales: hasHeader('Sales Name'),
    salesManager: hasHeader('Sales Manager'),
    nationality: hasHeader('Nationality') || hasHeader('Client Nationality'),
    clientType: hasHeader('Client Type'),

    // Branch aliases supported so Jeddah / Riyadh / Egypt branches
    // are picked up even if the sheet header name differs.
    branch:
      hasHeader('Branch') ||
      hasHeader('Branch Name') ||
      hasHeader('Sales Branch') ||
      hasHeader('Office') ||
      hasHeader('Office Branch'),

    status: hasHeader('Status'),
    renovation: hasHeader('Additional Status'),
    mainSource: sourceDetailsIdx['Main Source'] !== undefined || hasHeader('Main Source'),
    sourceType: sourceDetailsIdx['Source Type'] !== undefined || hasHeader('Source Type'),
    campaignName: sourceDetailsIdx['Campaign Name'] !== undefined || hasHeader('Campaign Name'),
    mediaBuyer: sourceDetailsIdx['Media Buyer'] !== undefined || hasHeader('Media Buyer'),
    brokerageCompany: sourceDetailsIdx['Brokerage Company'] !== undefined || hasHeader('Brokerage Company'),
    bcSales: sourceDetailsIdx['BC Sales'] !== undefined || hasHeader('BC Sales'),
    bcManager: sourceDetailsIdx['BC Manger'] !== undefined || hasHeader('BC Manger') || hasHeader('BC Manager'),
    day: hasHeader('Date'),
    month: hasHeader('Month') || hasHeader('Date'),
    year: hasHeader('Year') || hasHeader('Date')
  };

  const rows = [];

  // All branch names from the complete Transactions dataset.
  // Used only as FILTER OPTIONS for Egypt Viewer.
  // This does NOT expose rows/projects outside the Egypt Viewer permission scope.
  const allBranchesForFilter = Array.from(
    new Set(
      displayRows
        .map(r => String(
          getAny_(r, idx, [
            'Branch',
            'Branch Name',
            'Sales Branch',
            'Office',
            'Office Branch'
          ]) || ''
        ).trim())
        .filter(Boolean)
    )
  ).sort((a, b) =>
    String(a).localeCompare(String(b), undefined, {
      numeric: true,
      sensitivity: 'base'
    })
  );

  displayRows.forEach((r, i) => {
    const project = getAny_(r, idx, ['Project']);

    const clientName = cleanClientName_(
      getAny_(r, idx, ['Final Client Name', 'Client Name'])
    );

    const mobile = String(
      (
        clientInfoIdx['Contact'] !== undefined
          ? r[clientInfoIdx['Contact']]
          : getAny_(r, idx, [
              'Contact',
              'Mobile',
              'Mobile Number',
              'Client Mobile',
              'Client Mobile Number',
              'Phone',
              'Phone Number',
              'Client Phone'
            ])
      ) || ''
    ).trim();

    const nationality = String(
      getAny_(r, idx, [
        'Nationality',
        'Client Nationality',
        'Nationalty'
      ]) || ''
    ).trim();

    // SECURITY BOUNDARY:
    // Egypt Viewer never receives non-Egypt / non-EDEN WALK records.
    if (egyptViewer) {
      const projectKey = String(project || '').trim().toUpperCase();
      if (projectKey !== 'EDEN WALK') return;
      if (!isEgyptNationality_(nationality)) return;
    }

    // STABILITY V1: Client Type comes ONLY from the exact sheet header.
    const clientType = normalizeClientType_(
      getAny_(r, idx, ['Client Type'])
    );

    const unitCode = getAny_(r, idx, ['Unit Code']);

    if (!project && !clientName && !unitCode) return;

    const dealDate = getAny_(r, idx, ['Date']);
    const parsedDate = parseDate_(dealDate);
    const rawYear = getAny_(r, idx, ['Year']);
    const rawMonth = getAny_(r, idx, ['Month']);

    const dealYear =
      normalizeYear_(rawYear) ||
      (parsedDate ? parsedDate.getFullYear() : '');

    const dealMonthNo =
      normalizeMonth_(rawMonth) ||
      (parsedDate ? parsedDate.getMonth() + 1 : '');

    const dealMonthName = dealMonthNo
      ? monthName_(dealMonthNo)
      : '';

    const dealDay = parsedDate
      ? parsedDate.getDate()
      : '';

    const rawStatus = getAny_(r, idx, ['Status']);
    const additionalStatus = getAny_(r, idx, ['Additional Status']);

    const isRenovation =
      String(additionalStatus || '').trim().toLowerCase() === 'renovation';

    const status = rawStatus;

    const renovationStatus = isRenovation
      ? 'Renovation'
      : 'Without Renovation';

    const finalPrice = parseNumber_(
      getAny_(r, idx, ['Final Price'])
    );

    const totalPrice = parseNumber_(
      getAny_(r, idx, ['Primary Total Price'])
    );

    const amount =
      finalPrice ||
      totalPrice ||
      parseNumber_(getAny_(r, idx, ['Amount']));

    const currency = normalizeCurrency_(
      getAny_(r, idx, ['Currency'])
    );

    const area = parseNumber_(
      getAny_(r, idx, ['Area'])
    );

    // ===== DOWN PAYMENT DETAILS =====
    // Source of truth = exact header names in Transactions.
    // No column letters, no fixed indexes, no recalculation.
    const dpPercent = parseNumber_(
      getAny_(r, idx, ['DP %'])
    );

    const actualDP = parseNumber_(
      getAny_(r, idx, ['Actual DP', 'Acctual DP'])
    );

    const dpPaid = parseNumber_(
      getAny_(r, idx, ['DP Paid'])
    );

    const actualPaid = parseNumber_(
      getAny_(r, idx, ['Actual Paid', 'Acctual Paid'])
    );

    const remain = parseNumber_(
      getAny_(r, idx, ['Remain'])
    );

    const rawUnitType = String(
      (
        unitInfoIdx['Unit Type'] !== undefined
          ? r[unitInfoIdx['Unit Type']]
          : getAny_(r, idx, [
              'Unit Type',
              'Unit Category',
              'UnitType',
              'Unit_Type',
              'Type'
            ])
      ) || ''
    ).trim();

    const unitType = normalizeUnitType_(rawUnitType);
    const unitCategory = classifyUnitType_(unitType);

    // E-ONE: only Commercial / Admin / Medical units count as deals
    // (parking and anything else sold there is not a deal).
    if (isEOneProject_(project) && E_ONE_DEAL_UNIT_TYPES_.indexOf(unitType) < 0) return;

    rows.push({
      row: DATA_START_ROW + i,
      deal: getAny_(r, idx, ['Deal Num', 'Deal\nNum']) || i + 1,
      num: getAny_(r, idx, ['Deal Num', 'Deal\nNum']),

      year: dealYear,
      month: dealMonthNo,
      day: dealDay,
      monthName: dealMonthName,
      yearMonth:
        dealYear && dealMonthNo
          ? `${dealYear}-${String(dealMonthNo).padStart(2, '0')}`
          : '',

      project,
      clientName,
      clientKey: normalizeClientKey_(clientName),
      docsUrl: clientDocLinks[i] || '',
      sysAdded: systemAdded[i] ? 1 : 0,
      mobile,
      nationality,
      clientType,

      unitCode,
      status,
      additionalStatus,
      isRenovation,
      renovationStatus,
      date: dealDate,

      clientSignDate: getAny_(r, idx, ['Client Sign Date']),
      contractDate: contractDetailsIdx['Date'] !== undefined ? String(r[contractDetailsIdx['Date']] || '').trim() : '',
      companySignDate: getAny_(r, idx, ['Company Sign Date']),
      delivered: getAny_(r, idx, ['Delivered']),
      daysToClose: parseNumber_(
        getAny_(r, idx, ['Days to Close'])
      ),
      clientStatue: getAny_(r, idx, ['Client Statue']),

      block: getAny_(r, idx, ['Block']),
      view: getAny_(r, idx, ['View']),
      floor: getAny_(r, idx, ['Floor']),
      finishing: getAny_(r, idx, ['Finishing Type']),

      unitType,
      unitCategory,
      area,
      outArea: parseNumber_(
        getAny_(r, idx, ['OutArea'])
      ),
      outdoor: getAny_(r, idx, ['OutDoor']),

      meterPrice:
        parseNumber_(
          getAny_(r, idx, ['Meter Price After Discount'])
        ) ||
        parseNumber_(
          getAny_(r, idx, ['Primary Meter Price'])
        ),

      amount,
      currency,

      // Down Payment / Contract Deposit
      dpPercent,
      actualDP,
      dpPaid,
      actualPaid,
      remain,

      maintenanceAmount: parseNumber_(
        getAny_(r, idx, ['Maintenance Amount'])
      ),

      discountOffer: getAny_(r, idx, ['Discount Offer']),
      maintenancePercent: getAny_(r, idx, ['Maintenance %']),
      confPartDP: confIdx['Part of DP'] !== undefined ? r[confIdx['Part of DP']] : '',
      confTotalDP: confIdx['Total DP'] !== undefined ? r[confIdx['Total DP']] : '',
      confClientSigned: confIdx['Client Sigend'] !== undefined ? r[confIdx['Client Sigend']] : '',
      confCompanySigned: confIdx['Company Signed'] !== undefined ? r[confIdx['Company Signed']] : '',
      confDeliverySigned: confIdx['Delivery Signed'] !== undefined ? r[confIdx['Delivery Signed']] : '',
      confDelivered: confIdx['Delivered'] !== undefined ? r[confIdx['Delivered']] : '',
      dpDate: dpDetailsIdx['Date'] !== undefined ? String(r[dpDetailsIdx['Date']] || '').trim() : '',
      planType: planIdx['Type Of Payment Plan'] !== undefined ? String(r[planIdx['Type Of Payment Plan']] || '').trim() : '',
      planPeriod: planIdx['Installments Period'] !== undefined ? String(r[planIdx['Installments Period']] || '').trim() : '',
      planCode: planIdx['Installment Plan'] !== undefined ? String(r[planIdx['Installment Plan']] || '').trim() : '',
      planEvery: planIdx['Period Type'] !== undefined ? String(r[planIdx['Period Type']] || '').trim() : '',

      // Final branch mapping with aliases.
      branch:
        getAny_(r, idx, [
          'Branch',
          'Branch Name',
          'Sales Branch',
          'Office',
          'Office Branch'
        ]) || 'Not Assigned',

      salesName: getAny_(r, idx, ['Sales Name']) || 'Not Assigned',
      sharedWith: getAny_(r, idx, ['Shared With']),
      salesManager: getAny_(r, idx, ['Sales Manager']),

      // ===== INTERNAL COMMISSIONS =====
      // These are the four checkbox columns under the merged "Internal comm" group.
      // Google Sheets checkboxes are returned by getDisplayValues() as TRUE / FALSE.
      internalCommSales: internalCommIdx.Sales !== undefined ? r[internalCommIdx.Sales] : '',
      internalCommManagers: internalCommIdx.Managers !== undefined ? r[internalCommIdx.Managers] : '',
      internalCommStaff: internalCommIdx.Staff !== undefined ? r[internalCommIdx.Staff] : '',
      internalCommStatue: internalCommIdx.Statue !== undefined ? r[internalCommIdx.Statue] : '',

      headOfSales: getAny_(r, idx, ['Head Of Sales']),
      cco: getAny_(r, idx, ['CCO']),
      dealStatus: getAny_(r, idx, ['Deal Status']),

      mainSource: sourceDetailsIdx['Main Source'] !== undefined ? r[sourceDetailsIdx['Main Source']] : getAny_(r, idx, ['Main Source']),
      sourceType: sourceDetailsIdx['Source Type'] !== undefined ? r[sourceDetailsIdx['Source Type']] : getAny_(r, idx, ['Source Type']),
      campaignName: sourceDetailsIdx['Campaign Name'] !== undefined ? r[sourceDetailsIdx['Campaign Name']] : getAny_(r, idx, ['Campaign Name']),
      mediaBuyer: sourceDetailsIdx['Media Buyer'] !== undefined ? r[sourceDetailsIdx['Media Buyer']] : getAny_(r, idx, ['Media Buyer']),
      brokerageCompany: sourceDetailsIdx['Brokerage Company'] !== undefined ? r[sourceDetailsIdx['Brokerage Company']] : getAny_(r, idx, ['Brokerage Company']),
      bcSales: sourceDetailsIdx['BC Sales'] !== undefined ? r[sourceDetailsIdx['BC Sales']] : getAny_(r, idx, ['BC Sales']),
      bcManager: sourceDetailsIdx['BC Manger'] !== undefined ? r[sourceDetailsIdx['BC Manger']] : getAny_(r, idx, ['BC Manger', 'BC Manager'])
    });
  });

  const options = buildOptions_(rows, availableFilters);

  // SOURCE DETAILS filter options are built directly from mapped rows.
  const sourceOptionList_ = key => Array.from(
    new Set(rows.map(x => String(x[key] || '').trim()).filter(Boolean))
  ).sort((a,b) => String(a).localeCompare(String(b), undefined, {numeric:true,sensitivity:'base'}));

  options.mainSources = sourceOptionList_('mainSource');
  options.sourceTypes = sourceOptionList_('sourceType');
  options.campaignNames = sourceOptionList_('campaignName');
  options.mediaBuyers = sourceOptionList_('mediaBuyer');
  options.brokerageCompanies = sourceOptionList_('brokerageCompany');
  options.bcSales = sourceOptionList_('bcSales');
  options.bcManagers = sourceOptionList_('bcManager');

  // UNIT TYPE: source options directly from the mapped transaction rows.
  // This bypasses buildOptions_ alias differences completely.
  const mappedUnitTypes = Array.from(
    new Set(
      rows
        .map(x => String(x.unitType || '').trim())
        .filter(Boolean)
    )
  ).sort((a,b) => String(a).localeCompare(String(b), undefined, {
    numeric: true,
    sensitivity: 'base'
  }));

  options.unitTypes = mappedUnitTypes;
  options.unitCategories = mappedUnitTypes.slice();
  options.unitCategory = mappedUnitTypes.slice();
  options.salesManagers = availableFilters.salesManager
    ? uniqueFilterValues_(rows, 'salesManager', isValidSalesFilterValue_)
    : [];

  // Egypt Viewer:
  // - Project options stay restricted to EDEN WALK because they come from `rows`.
  // - Branch filter shows ALL company branches from the complete Transactions sheet.
  // - Selecting a branch with no permitted EDEN WALK/Egyptian rows simply returns 0 rows.
  if (egyptViewer) {
    const wantedBranchKeys = new Set([
      'EGY - AL-REHAB',
      'EGY - AL-REHAB BRANCH',
      'KSA - JEDDAH',
      'KSA - JEDDAH BRANCH',
      'KSA - RIYADH',
      'KSA - RIYADH BRANCH'
    ]);

    options.branches = allBranchesForFilter.filter(branch => {
      const key = String(branch || '')
        .trim()
        .toUpperCase()
        .replace(/\s+/g, ' ');
      return wantedBranchKeys.has(key);
    });
  }

  return {
    sheetName: sheet.getName(),
    generatedAt: new Date().toISOString(),
    rows,
    options,
    brokers: getBrokerDashboardData_(),   // broker companies are visible to Egypt Viewer too
    access: egyptViewerAccessProfile_(authUser)
  };
}


const E_ONE_DEAL_UNIT_TYPES_ = ['Commercial', 'Admin', 'Medical'];
function isEOneProject_(project) {
  return String(project || '').toUpperCase().replace(/[^A-Z]/g, '') === 'EONE';
}

/* true for each data row whose Project cell has the EDEN BI note. Never throws. */
function dealsSystemNotes_(sheet, headers, rowCount) {
  const out = new Array(rowCount).fill(false);
  try {
    const c = headers.map(h => normalizeHeader_(h)).indexOf('project');
    if (c < 0 || !rowCount) return out;
    sheet.getRange(DATA_START_ROW, c + 1, rowCount, 1).getNotes()
      .forEach((n, i) => { out[i] = /^EDEN BI/.test(String(n[0] || '')); });
  } catch (err) {}
  return out;
}

/**
 * Links on the client name cells (the client's documents folder / file).
 * Reads both inserted links and =HYPERLINK() formulas. Returns one URL (or '')
 * per data row; never throws, so the dashboard still loads without them.
 */
function clientDocumentLinks_(sheet, idx, rowCount) {
  const out = new Array(rowCount).fill('');
  if (!rowCount) return out;
  const cols = ['Client Name', 'Final Client Name']
    .map(n => idx[n] !== undefined ? idx[n] : idx[normalizeHeader_(n)])
    .filter((c, i, a) => c !== undefined && a.indexOf(c) === i);
  cols.forEach(c => {
    try {
      const range = sheet.getRange(DATA_START_ROW, c + 1, rowCount, 1);
      const rich = range.getRichTextValues();
      const formulas = range.getFormulas();
      for (let i = 0; i < rowCount; i++) {
        if (out[i]) continue;
        let url = '';
        const rt = rich[i] && rich[i][0];
        if (rt) {
          url = rt.getLinkUrl() || '';
          if (!url) {
            const runs = rt.getRuns();
            for (let k = 0; k < runs.length && !url; k++) url = runs[k].getLinkUrl() || '';
          }
        }
        if (!url) {
          const m = String((formulas[i] && formulas[i][0]) || '').match(/HYPERLINK\(\s*"([^"]+)"/i);
          if (m) url = m[1];
        }
        if (/^https?:\/\//i.test(url)) out[i] = url;
      }
    } catch (err) {
      console.warn('Client document links: ' + err);
    }
  });
  return out;
}
