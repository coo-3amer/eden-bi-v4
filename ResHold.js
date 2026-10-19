/************************************************
 * EDEN BI V4 — ResHold.js
 *
 * A reservation is held for 7 days at most. When a NEW reservation is still
 * "Reservation" 7 days after its Reservation Date, the deal is cancelled in
 * the Deals sheet (Cancel Details: Check ✓, Reason, Date — the Status formula
 * turns it to Cancel) and the unit goes back to Available in the Inventory.
 *   - "New" = Reservation Date on/after the day this feature started
 *     (RES_HOLD_SINCE) and not one of the reservations already open then
 *     (the first run remembers them in RES_HOLD_OLD). Old ones keep the old rules.
 *   - One day before the end, a "Reservation expiring" alert is sent.
 *   - Runs from the 5-minute warm-up, at most every 10 minutes.
 *   - Script Property RES_HOLD_MINUTES (optional) overrides the 7 days.
 *   - Test: a deal whose Unit Code contains "TEST" is held 10 minutes only.
 ************************************************/

const RES_HOLD_DAYS_ = 7;
const RES_HOLD_REASON_ = 'انتهت مدة الحجز (7 أيام) بدون سداد المقدم — إلغاء تلقائي';

function resHoldKey_(code, unit) {
  return (String(code || '').trim() + '|' + String(unit || '').trim()).toLowerCase();
}

/* Reserved before the day the 7-day rule started (Script Property RES_HOLD_SINCE). */
function resHoldIsOldDate_(d) {
  const since = PropertiesService.getScriptProperties().getProperty('RES_HOLD_SINCE');
  if (!since || !(d instanceof Date)) return false;
  const s = new Date(since);
  return d.getTime() < new Date(s.getFullYear(), s.getMonth(), s.getDate()).getTime();
}

function resHoldMs_() {
  const m = Number(PropertiesService.getScriptProperties().getProperty('RES_HOLD_MINUTES'));
  return (m > 0 ? m : RES_HOLD_DAYS_ * 1440) * 60000;
}

function resHoldColumns_(sh) {
  const cols = mainDealsColumns_(sh);
  const at = names => mainDealsCol_(cols, names);
  const rsn = at(['cancel details|reason']), cdt = at(['cancel details|date']);
  // The cancel "Check" box sits just before the Cancel Details group (under the
  // previous group title), so take the nearest "Check" header left of Reason.
  let chk = at(['cancel details|check']);
  if (!chk && rsn) {
    const heads = sh.getRange(HEADER_ROW, 1, 1, rsn).getDisplayValues()[0];
    for (let k = rsn - 1; k >= Math.max(1, rsn - 4); k--) if (normalizeHeader_(heads[k - 1]) === normalizeHeader_('Check')) { chk = k; break; }
  }
  return {
    cols: cols,
    p: at(['Project']), u: at(['Unit Code']), s: at(['deal status|status', 'Status']),
    code: at(['reservation details|res code', 'Res Code']), date: at(['reservation details|date']),
    cl: at(['Final Client Name', 'Client Name']),
    chk: chk, rsn: rsn, cdt: cdt
  };
}

/* Called by warmDashboardCache. Never throws. */
function reservationHoldJob_(force) {
  try {
    const cache = CacheService.getScriptCache();
    if (!force && cache.get('RES_HOLD_AT')) return;
    cache.put('RES_HOLD_AT', '1', 590);
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return;
    try { return resHoldRun_(); } finally { lock.releaseLock(); }
  } catch (err) {
    console.warn('reservationHoldJob_: ' + err);
  }
}

function resHoldRun_() {
  const props = PropertiesService.getScriptProperties();
  const sh = getMainDealsSheet_();
  const c = resHoldColumns_(sh);
  if (!c.p || !c.u || !c.s || !c.date || !c.chk || !c.cdt) {
    console.warn('Reservation hold: Deals columns not found ' + JSON.stringify(c, (k, v) => k === 'cols' ? undefined : v));
    return { error: 'columns' };
  }
  const last = mainDealsLastRow_(sh, c.cols);
  if (last < DATA_START_ROW) return { done: 0 };
  const n = last - DATA_START_ROW + 1;
  const disp = k => k ? sh.getRange(DATA_START_ROW, k, n, 1).getDisplayValues().map(r => String(r[0] || '').trim()) : new Array(n).fill('');
  const P = disp(c.p), U = disp(c.u), S = disp(c.s), K = disp(c.code), CL = disp(c.cl);
  const D = sh.getRange(DATA_START_ROW, c.date, n, 1).getValues().map(r => r[0]);
  const isRes = i => S[i].toLowerCase() === 'reservation';

  // First run: every reservation open now keeps the old rules.
  let old = props.getProperty('RES_HOLD_OLD');
  if (old === null) {
    const keys = [];
    for (let i = 0; i < n; i++) if (isRes(i) && U[i]) keys.push(resHoldKey_(K[i], U[i]));
    props.setProperties({ RES_HOLD_OLD: JSON.stringify(keys), RES_HOLD_SINCE: new Date().toISOString() });
    auditLog_('System', 'Reservation 7-day limit on', { details: 'Applies to new reservations only. ' + keys.length + ' reservations already open keep the old rules.' });
    return { started: keys.length };
  }
  const oldSet = {};
  JSON.parse(old || '[]').forEach(k => oldSet[k] = 1);

  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const fmt = d => Utilities.formatDate(d, tz, 'dd-MMM-yyyy');
  const hold = resHoldMs_(), now = Date.now();
  const warned = JSON.parse(props.getProperty('RES_HOLD_WARNED') || '{}');
  const out = { expired: [], warned: [], tracked: 0 };

  for (let i = 0; i < n; i++) {
    if (!isRes(i) || !U[i]) continue;
    const key = resHoldKey_(K[i], U[i]);
    if (oldSet[key]) continue;
    const d = D[i] instanceof Date ? D[i] : null;
    if (!d) continue;                                                 // no Reservation Date: days can't be counted
    if (resHoldIsOldDate_(d)) continue;                               // reserved before the 7-day rule started
    out.tracked++;
    // Test units (code with "TEST") are held 10 minutes from when the system first sees them.
    const test = /TEST/i.test(U[i]);
    if (test && !warned['seen:' + key]) warned['seen:' + key] = now;
    const start = test ? warned['seen:' + key] : new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();   // reservation day, 00:00
    const end = start + (test ? 10 * 60000 : hold);
    const row = DATA_START_ROW + i;
    const info = { project: P[i], unit: U[i], client: CL[i], ref: 'Deals row ' + row };

    if (now < end) {
      if (end - now <= Math.min(24 * 3600000, hold / 2) && !warned[key]) {
        warned[key] = now;
        auditLog_('System', 'Reservation expiring', Object.assign({}, info, {
          details: 'Reserved ' + fmt(d) + ' — the reservation ends ' + Utilities.formatDate(new Date(end), tz, 'dd-MMM HH:mm') +
            '. Without a down payment by then the deal is cancelled and the unit is opened.' }));
        out.warned.push(U[i]);
      }
      continue;
    }

    // Expired: check the live row once more, then cancel it.
    const live = sh.getRange(row, 1, 1, c.cols.lastCol).getDisplayValues()[0];
    if (String(live[c.s - 1] || '').trim().toLowerCase() !== 'reservation' || String(live[c.u - 1] || '').trim() !== U[i]) continue;
    const today = new Date();
    sh.getRange(row, c.chk).setValue(true);
    if (c.rsn && !String(live[c.rsn - 1] || '').trim()) sh.getRange(row, c.rsn).setValue(RES_HOLD_REASON_);
    sh.getRange(row, c.cdt).setValue(new Date(today.getFullYear(), today.getMonth(), today.getDate())).setNumberFormat('d-mmm-yyyy');
    SpreadsheetApp.flush();
    const ns = String(sh.getRange(row, c.s).getDisplayValue() || '').trim() || 'Cancel';
    if (typeof statusSeenRemember_ === 'function') statusSeenRemember_(row, ns);
    const inv = syncInventoryForDeal_(P[i], U[i], ns, { by: 'reservation expired' });
    delete warned[key]; delete warned['seen:' + key];
    auditLog_('System', 'Reservation expired', Object.assign({}, info, {
      details: 'Reserved ' + fmt(d) + ' — ' + (test ? '10 minutes (test unit)' : Math.round(hold / 86400000 * 10) / 10 + ' days') + ' passed without a down payment. Deal → ' + ns +
        (inv && inv.message ? ' · ' + inv.message : '') }));
    out.expired.push(U[i]);
  }
  // forget warnings of deals that are no longer reservations
  Object.keys(warned).forEach(k => { if (now - warned[k] > 30 * 86400000) delete warned[k]; });
  props.setProperty('RES_HOLD_WARNED', JSON.stringify(warned));
  return out;
}

function resHoldAdmin_(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Only admins can see and cancel reservations.');
  return user;
}

/* Admins: every open reservation (Transactions › Reservations page). */
function getReservationHolds(authToken) {
  resHoldAdmin_(authToken);
  const props = PropertiesService.getScriptProperties();
  const sh = getMainDealsSheet_();
  const c = resHoldColumns_(sh);
  const at = names => mainDealsCol_(c.cols, names);
  const sc = at(['Sales Name']), am = at(['reservation details|amount']), cu = at(['reservation details|currency']);
  const last = mainDealsLastRow_(sh, c.cols);
  const oldSet = {};
  JSON.parse(props.getProperty('RES_HOLD_OLD') || '[]').forEach(k => oldSet[k] = 1);
  const hold = resHoldMs_();
  const out = { since: props.getProperty('RES_HOLD_SINCE') || '', holdDays: Math.round(hold / 86400000 * 100) / 100, rows: [] };
  if (last < DATA_START_ROW || !c.s || !c.u) return out;
  const n = last - DATA_START_ROW + 1;
  const disp = k => k ? sh.getRange(DATA_START_ROW, k, n, 1).getDisplayValues().map(r => String(r[0] || '').trim()) : new Array(n).fill('');
  const P = disp(c.p), U = disp(c.u), S = disp(c.s), K = disp(c.code), CL = disp(c.cl), SA = disp(sc), AM = disp(am), CU = disp(cu);
  const D = c.date ? sh.getRange(DATA_START_ROW, c.date, n, 1).getValues().map(r => r[0]) : new Array(n).fill('');
  for (let i = 0; i < n; i++) {
    if (S[i].toLowerCase() !== 'reservation' || !U[i]) continue;
    const d = D[i] instanceof Date ? D[i] : null;
    const isOld = !!oldSet[resHoldKey_(K[i], U[i])] || resHoldIsOldDate_(d);
    const end = d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() + hold : 0;
    out.rows.push({ row: DATA_START_ROW + i, project: P[i], unit: U[i], client: CL[i], resCode: K[i], sales: SA[i],
      amount: AM[i], currency: CU[i], resDate: d ? d.toISOString() : '', old: isOld, endsAt: (!isOld && end) ? new Date(end).toISOString() : '' });
  }
  return out;
}

/* Admins: cancel one reservation now — Deals (Check ✓, Reason, Date → Status Cancel) + unit back to Available. */
function cancelReservation(authToken, req) {
  const user = resHoldAdmin_(authToken);
  req = req || {};
  const row = Number(req.row), unit = String(req.unit || '').trim(), reason = String(req.reason || '').trim();
  if (!row || row < DATA_START_ROW || !unit) throw new Error('Choose a reservation.');
  if (!reason) throw new Error('Write the reason for cancelling.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The system is busy, try again in a few seconds.');
  try {
    const sh = getMainDealsSheet_();
    const c = resHoldColumns_(sh);
    if (!c.s || !c.u || !c.chk || !c.rsn || !c.cdt) throw new Error('Cancel columns not found in the Deals sheet.');
    const live = sh.getRange(row, 1, 1, c.cols.lastCol).getDisplayValues()[0];
    const liveUnit = String(live[c.u - 1] || '').trim(), status = String(live[c.s - 1] || '').trim();
    if (liveUnit !== unit) throw new Error('The Deals sheet changed (row ' + row + ' is now ' + (liveUnit || 'empty') + '). Reload the list and try again.');
    if (status.toLowerCase() !== 'reservation') throw new Error(unit + ' is not a reservation any more (status: ' + (status || '—') + ').');
    const name = (user && (user.name || user.username)) || 'Admin';
    const today = new Date();
    sh.getRange(row, c.chk).setValue(true);
    sh.getRange(row, c.rsn).setValue(reason);
    sh.getRange(row, c.cdt).setValue(new Date(today.getFullYear(), today.getMonth(), today.getDate())).setNumberFormat('d-mmm-yyyy');
    SpreadsheetApp.flush();
    const ns = String(sh.getRange(row, c.s).getDisplayValue() || '').trim() || 'Cancel';
    if (typeof statusSeenRemember_ === 'function') statusSeenRemember_(row, ns);
    const project = String(live[c.p - 1] || '').trim(), client = c.cl ? String(live[c.cl - 1] || '').trim() : '';
    const inv = syncInventoryForDeal_(project, unit, ns, { by: name });
    auditLog_(user, 'Reservation cancelled', { project: project, unit: unit, client: client,
      details: 'Reason: ' + reason + ' · Deal → ' + ns + (inv && inv.message ? ' · ' + inv.message : ''), ref: 'Deals row ' + row });
    try { clearDashboardCache_(); } catch (e) {}
    return { ok: true, status: ns, inventory: inv || {}, message: unit + ' cancelled' + (inv && inv.message ? ' · ' + inv.message : '') };
  } finally {
    lock.releaseLock();
  }
}
