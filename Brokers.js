function getBrokerDashboardData_() {
  const ss = SpreadsheetApp.openById(BROKER_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(BROKER_SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) return { sheetName: '', rows: [], options: {}, availableFields: {} };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < BROKER_DATA_START_ROW) {
    return { sheetName: sheet.getName(), rows: [], options: {}, availableFields: {} };
  }

  const headers = sheet
    .getRange(BROKER_HEADER_ROW, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(h => String(h || '').trim());

  const displayRows = sheet
    .getRange(BROKER_DATA_START_ROW, 1, lastRow - BROKER_DATA_START_ROW + 1, lastCol)
    .getDisplayValues();

  const idx = buildIndex_(headers);
  const rows = [];

  // Detect whether a field actually exists in the source sheet by header name.
  const hasAnyHeader_ = aliases => aliases.some(name =>
    idx[name] !== undefined ||
    idx[normalizeHeader_(name)] !== undefined
  );

  const fieldAliases = {
    company: ['Company', 'Broker Company', 'Company Name'],
    creationDate: ['Creation Date', 'Created Date', 'Date Created'],
    status: ['Status', 'Company Status'],
    edenSales: ['Eden Sales', 'EDEN Sales'],
    firstSales: ['First Sales', '1st Sales'],
    secondSales: ['Second Sales', '2nd Sales'],
    lastSales: ['Last Sales', 'Latest Sales'],
    createdBy: ['Created By', 'Creator', 'Added By'],
    admin: ['Admin', 'Admin Name', 'Company Admin'],
    mobile: ['Admin Mobile', 'Admin Mobile Number', 'Admin Phone', 'Admin Phone Number', 'Mobile'],
    brokerSales: ['Sales', 'Broker Sales', 'Sales Count'],
    groupLink: ['Group Link', 'WhatsApp Group Link', 'Whatsapp Group Link', 'Group URL'],
    orientation1: ['Oriantation 1', 'Orientation 1'],
    orientation2: ['Oriantation 2', 'Orientation 2'],
    orientation3: ['Oriantation 3', 'Orientation 3'],
    orientation4: ['Oriantation 4', 'Orientation 4'],
    orientation5: ['Oriantation 5', 'Orientation 5'],
    workshop1: ['Workshop 1'],
    workshop2: ['Workshop 2']
  };

  const availableFields = Object.fromEntries(
    Object.entries(fieldAliases).map(([key, aliases]) => [key, hasAnyHeader_(aliases)])
  );

  displayRows.forEach((r, i) => {
    const company = getAny_(r, idx, fieldAliases.company);
    if (!company) return;

    const creationDate = getAny_(r, idx, fieldAliases.creationDate);
    const parsedDate = parseDate_(creationDate);
    const year = parsedDate ? parsedDate.getFullYear() : '';
    const month = parsedDate ? parsedDate.getMonth() + 1 : '';
    const day = parsedDate ? parsedDate.getDate() : '';

    const directEdenSales = getAny_(r, idx, fieldAliases.edenSales);
    const firstSales = getAny_(r, idx, fieldAliases.firstSales);
    const secondSales = getAny_(r, idx, fieldAliases.secondSales);
    const lastSales = getAny_(r, idx, fieldAliases.lastSales);

    const salesPeople = [...new Set([
      directEdenSales,
      firstSales,
      secondSales,
      lastSales
    ].map(v => String(v || '').trim()).filter(v => v && v !== 'Not Assigned'))];

    // The company's CURRENT EDEN salesperson is the "Eden Sales" column only.
    // First / Second / Last Sales are history and must not stand in for it.
    const primaryEdenSales = availableFields.edenSales
      ? (String(directEdenSales || '').trim() || 'Not Assigned')
      : (lastSales || secondSales || firstSales || 'Not Assigned');

    const groupLink = getAny_(r, idx, fieldAliases.groupLink);
    const adminValue = getAny_(r, idx, fieldAliases.admin);
    const mobileValue = getAny_(r, idx, fieldAliases.mobile);

    rows.push({
      row: BROKER_DATA_START_ROW + i,
      company,
      companyKey: normalizeClientKey_(company),
      creationDate,
      year,
      month,
      day,
      monthName: month ? monthName_(month) : '',
      yearMonth: year && month ? `${year}-${String(month).padStart(2, '0')}` : '',
      status: getAny_(r, idx, fieldAliases.status) || 'Not Specified',
      firstSales,
      secondSales,
      lastSales,
      salesPeople,
      createdBy: getAny_(r, idx, fieldAliases.createdBy) || 'Not Assigned',
      edenSales: primaryEdenSales,

      // If the header is absent, keep the value empty.
      // Company Health will ignore this field completely.
      admin: availableFields.admin ? adminValue : '',
      mobile: availableFields.mobile ? mobileValue : '',

      brokerSales: parseNumber_(getAny_(r, idx, fieldAliases.brokerSales)),
      groupLink,
      linkStatus: groupLink ? 'with' : 'without',
      linkStatusLabel: groupLink ? 'Company With Link' : 'Company Without Link',
      orientation1: getAny_(r, idx, fieldAliases.orientation1),
      orientation2: getAny_(r, idx, fieldAliases.orientation2),
      orientation3: getAny_(r, idx, fieldAliases.orientation3),
      orientation4: getAny_(r, idx, fieldAliases.orientation4),
      orientation5: getAny_(r, idx, fieldAliases.orientation5),
      workshop1: getAny_(r, idx, fieldAliases.workshop1),
      workshop2: getAny_(r, idx, fieldAliases.workshop2)
    });
  });

  return {
    sheetName: sheet.getName(),
    rows,
    options: buildBrokerOptions_(rows),
    availableFields
  };
}

function buildBrokerOptions_(rows) {
  const unique = key => [...new Set(
    rows.map(r => r[key]).filter(v =>
      v !== '' &&
      v !== null &&
      v !== undefined &&
      v !== 'Not Assigned'
    )
  )].sort((a, b) =>
    String(a).localeCompare(String(b), undefined, { numeric: true })
  );

  return {
    companies: unique('company'),
    statuses: unique('status'),
    edenSales: [...new Set(
      rows.flatMap(r => r.salesPeople || []).filter(Boolean)
    )].sort((a, b) =>
      String(a).localeCompare(String(b), undefined, { numeric: true })
    ),
    admins: unique('admin'),
    createdBy: unique('createdBy'),
    lastSales: unique('lastSales'),
    days: unique('day'),
    months: unique('month').map(m => ({ value: m, label: monthName_(m) })),
    years: unique('year')
  };
}
