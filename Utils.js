function getSheetByGid_(ss, gid) {
  return ss.getSheets().find(s => String(s.getSheetId()) === String(gid));
}

function buildIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    if (!h) return;
    const clean = normalizeHeader_(h);
    if (idx[clean] === undefined) idx[clean] = i;
    if (idx[h] === undefined) idx[h] = i;
  });
  return idx;
}

function normalizeHeader_(h) {
  return String(h || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function getAny_(row, idx, names) {
  for (const name of names) {
    const direct = idx[name];
    if (direct !== undefined) return String(row[direct] || '').trim();
    const clean = idx[normalizeHeader_(name)];
    if (clean !== undefined) return String(row[clean] || '').trim();
  }
  return '';
}

function parseNumber_(v) {
  if (v === null || v === undefined) return 0;
  const s = String(v).replace(/[,£$€%]|EGP|USD|m²|متر/gi, '').replace(/\s+/g, '').trim();
  const n = Number(s.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function normalizeCurrency_(v) {
  const s = String(v || '').toUpperCase();
  if (s.includes('USD') || s.includes('$')) return 'USD';
  if (s.includes('EGP') || s.includes('£')) return 'EGP';
  return s || 'EGP';
}

function cleanClientName_(v) {
  return String(v || '')
    .replace(/\s+/g, ' ')
    .replace(/\s+-\s+.*$/g, '')
    .trim();
}

function normalizeClientKey_(v) {
  return String(v || '')
    .toLowerCase()
    .replace(/[\u064B-\u065F\u0670]/g, '')
    .replace(/[إأآا]/g, 'ا')
    .replace(/[ى]/g, 'ي')
    .replace(/[ة]/g, 'ه')
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseDate_(v) {
  if (!v) return null;
  const s = String(v).trim();

  // Google Sheets sometimes returns Excel serial dates as display values.
  const serial = Number(s);
  if (!isNaN(serial) && serial > 20000 && serial < 70000) {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + serial * 24 * 60 * 60 * 1000);
    if (!isNaN(d)) return d;
  }

  let d = new Date(s);
  if (!isNaN(d)) return d;
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const day = Number(m[1]);
    const month = Number(m[2]) - 1;
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    d = new Date(year, month, day);
    if (!isNaN(d)) return d;
  }
  return null;
}

function normalizeYear_(v) {
  const m = String(v || '').match(/20\d{2}/);
  return m ? Number(m[0]) : '';
}

function normalizeMonth_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const n = Number(s);
  if (n >= 1 && n <= 12) return n;
  const map = {
    jan:1,january:1,feb:2,february:2,mar:3,march:3,apr:4,april:4,may:5,jun:6,june:6,
    jul:7,july:7,aug:8,august:8,sep:9,september:9,oct:10,october:10,nov:11,november:11,dec:12,december:12
  };
  const key = s.toLowerCase();
  return map[key] || '';
}

function monthName_(m) {
  const names = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return names[Number(m) - 1] || '';
}

function cleanFilterValue_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/\s+/g, ' ')
    .trim();
}

function isGenericInvalidFilterValue_(value) {
  const v = cleanFilterValue_(value).toLowerCase();
  return !v || [
    '-', '--', '---',
    'n/a', 'na', 'none', 'null', 'undefined',
    'unknown', 'not assigned', 'not available'
  ].includes(v);
}

function isValidSalesFilterValue_(value) {
  const v = cleanFilterValue_(value);
  if (isGenericInvalidFilterValue_(v)) return false;

  // Values below are statuses/flags, not sales-person names.
  const blocked = [
    'not signed', 'signed', 'client signed', 'company signed',
    'delivered', 'closed', 'cancel', 'cancelled',
    'reservation', 'in progress', 'part of dp', 'total dp'
  ];

  return !blocked.includes(v.toLowerCase());
}

function uniqueFilterValues_(rows, key, validator) {
  const seen = new Map();

  rows.forEach(row => {
    const value = cleanFilterValue_(row[key]);
    if (isGenericInvalidFilterValue_(value)) return;
    if (validator && !validator(value)) return;

    const identity = value.toLowerCase();
    if (!seen.has(identity)) seen.set(identity, value);
  });

  return [...seen.values()].sort((a, b) =>
    String(a).localeCompare(String(b), undefined, {
      numeric: true,
      sensitivity: 'base'
    })
  );
}

function buildOptions_(rows, availableFilters) {
  const enabled = availableFilters || {};
  const values = (key, validator) => uniqueFilterValues_(rows, key, validator);

  return {
    availableFilters: enabled,
    projects: enabled.project ? values('project') : [],
    sales: enabled.sales ? values('salesName', isValidSalesFilterValue_) : [],
    branches: enabled.branch ? values('branch') : [],
    statuses: enabled.status ? values('status') : [],
    days: enabled.day ? values('day') : [],
    months: enabled.month
      ? values('month').map(m => ({ value: m, label: monthName_(m) }))
      : [],
    years: enabled.year ? values('year') : [],
    currencies: values('currency'),
    unitTypes: enabled.unitCategory ? sortUnitTypes_(values('unitType')) : [],
    unitCategories: enabled.unitCategory ? values('unitCategory') : [],
    sources: enabled.mainSource ? values('mainSource') : [],
    mainSources: enabled.mainSource ? values('mainSource') : [],
    sourceTypes: enabled.sourceType ? values('sourceType') : [],
    campaignNames: enabled.campaignName ? values('campaignName') : [],
    mediaBuyers: enabled.mediaBuyer ? values('mediaBuyer') : []
  };
}

function allowedUnitTypes_() {
  return [
    'Commercial',
    'Admin',
    'Medical',
    'Residential',
    'Hotel Apartment',
    'Studio',
    'One Bedroom',
    '2 Bedroom',
    '3 Bedroom',
    'Villa',
    'Ground Villa',
    'Roof Villa',
    'Sky Villa',
    'Sky Roof Villa'
  ];
}

function sortUnitTypes_(arr) {
  const order = allowedUnitTypes_();
  return arr
    .filter(x => x && x !== 'Not Classified')
    .sort((a, b) => {
      const ia = order.indexOf(a);
      const ib = order.indexOf(b);
      if (ia !== -1 || ib !== -1) return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
      return String(a).localeCompare(String(b), undefined, { numeric: true });
    });
}

function normalizeUnitType_(unitType) {
  const raw = String(unitType || '').replace(/\s+/g, ' ').trim();
  const t = raw.toLowerCase();
  if (!t) return '';

  if (t === 'commercial' || t.includes('commercial') || t.includes('retail') || t.includes('shop') || t.includes('store')) return 'Commercial';
  if (t === 'admin' || t.includes('admin') || t.includes('administrative') || t.includes('office')) return 'Admin';
  if (t === 'medical' || t.includes('medical') || t.includes('clinic') || t.includes('pharmacy')) return 'Medical';
  if (t === 'residential' || t.includes('residential')) return 'Residential';
  if (t === 'hotel apartment' || t.includes('hotel apartment') || t.includes('hotel-apartment') || t.includes('serviced apartment')) return 'Hotel Apartment';

  if (t === 'studio' || t.includes('studio')) return 'Studio';
  if (t === 'one bedroom' || t === '1 bedroom' || t === '1 bed' || t === '1br' || t === 'one bed') return 'One Bedroom';
  if (t === '2 bedroom' || t === 'two bedroom' || t === '2 bedrooms' || t === '2 bed' || t === '2br') return '2 Bedroom';
  if (t === '3 bedroom' || t === 'three bedroom' || t === '3 bedrooms' || t === '3 bed' || t === '3br') return '3 Bedroom';

  if (t === 'ground villa' || t.includes('ground villa')) return 'Ground Villa';
  if (t === 'roof villa' || t.includes('roof villa')) return 'Roof Villa';
  if (t === 'sky roof villa' || t.includes('sky roof villa')) return 'Sky Roof Villa';
  if (t === 'sky villa' || t.includes('sky villa')) return 'Sky Villa';
  if (t === 'villa' || t.includes('villa')) return 'Villa';

  return raw;
}

function classifyUnitType_(unitType) {
  const t = String(unitType || '').trim().toLowerCase();
  if (!t) return 'Not Classified';

  if (t.includes('hotel apartment') || t.includes('hotel-apartment') || t.includes('serviced apartment') || t.includes('hotel') || t.includes('suite')) {
    return 'Hotel Apartment';
  }
  if (t.includes('medical') || t.includes('clinic') || t.includes('pharmacy')) {
    return 'Medical';
  }
  if (t.includes('admin') || t.includes('administrative') || t.includes('office')) {
    return 'Admin';
  }
  if (t.includes('commercial') || t.includes('retail') || t.includes('shop') || t.includes('store')) {
    return 'Commercial';
  }
  if (t.includes('residential') || t.includes('apartment') || t.includes('studio') || t.includes('villa') || t.includes('chalet') || t.includes('duplex') || t.includes('penthouse')) {
    return 'Residential';
  }

  return unitType;
}

function logDashboardActivity(action, section, details) {
  return { ok: true };
}
