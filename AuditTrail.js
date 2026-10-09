/************************************************
 * EDEN BI V4 — AuditLog.js
 *
 * Activity log: every change made from the system is written to the
 * "Audit Log" tab of the main spreadsheet (one row per action):
 * deals added / updated, payment schedules, payments recorded or
 * deleted, Inventory status changes, sign-ins.
 * Writing never blocks the action itself. Only admins can read it.
 ************************************************/

const AUDIT_SHEET_ = 'Audit Log';
const AUDIT_HEAD_ = ['Time', 'User', 'Role', 'Action', 'Project', 'Unit', 'Client', 'Details', 'Ref'];

function auditSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(AUDIT_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(AUDIT_SHEET_);
    sh.getRange(1, 1, 1, AUDIT_HEAD_.length).setValues([AUDIT_HEAD_]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(8, 420);
  }
  return sh;
}

/* user: the signed-in user object, or just a name. Never throws. */
function auditLog_(user, action, info) {
  try {
    info = info || {};
    const name = typeof user === 'string' ? user : (user && (user.name || user.username)) || '';
    const role = typeof user === 'string' ? '' : (user && user.role) || '';
    const clip = v => String(v == null ? '' : v).slice(0, 2000);
    const sh = auditSheet_();
    sh.appendRow([new Date(), clip(name), clip(role), clip(action), clip(info.project), clip(info.unit), clip(info.client), clip(info.details), clip(info.ref)]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    console.warn('auditLog_: ' + e);
  }
}

/* Admins: the latest entries, newest first. */
function getAuditLog(authToken, req) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Only admins can see the activity log.');
  req = req || {};
  const limit = Math.max(50, Math.min(5000, Number(req.limit) || 2000));
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(AUDIT_SHEET_);
  if (!sh || sh.getLastRow() < 2) return { rows: [], total: 0 };
  const total = sh.getLastRow() - 1;
  const n = Math.min(limit, total), start = sh.getLastRow() - n + 1;
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const rows = sh.getRange(start, 1, n, AUDIT_HEAD_.length).getValues().reverse().map(r => ({
    time: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm') : String(r[0] || ''),
    user: String(r[1] || ''), role: String(r[2] || ''), action: String(r[3] || ''), project: String(r[4] || ''),
    unit: String(r[5] || ''), client: String(r[6] || ''), details: String(r[7] || ''), ref: String(r[8] || '')
  }));
  return { rows: rows, total: total };
}

/* Hand edits in the Deals sheet (called by the LiveSync edit trigger). */
function auditDealsSheetEdit_(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getSheetId() !== DEALS_GID) return;
    const r0 = e.range.getRow(), c0 = e.range.getColumn(), nR = e.range.getNumRows(), nC = e.range.getNumColumns();
    if (r0 + nR - 1 < DATA_START_ROW) return;                       // header rows
    let who = '';
    try { who = (e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (err) {}
    const cols = mainDealsColumns_(sheet);
    const heads = sheet.getRange(HEADER_ROW, 1, 1, cols.lastCol).getDisplayValues()[0];
    const groups = HEADER_ROW > 1 ? sheet.getRange(HEADER_ROW - 1, 1, 1, cols.lastCol).getDisplayValues()[0] : [];
    const groupOf = c => { for (let i = c - 1; i >= 0; i--) if (String(groups[i] || '').trim()) return String(groups[i]).trim(); return ''; };
    const colName = c => {
      const h = String(heads[c - 1] || '').replace(/\s+/g, ' ').trim() || ('Column ' + c);
      const g = groupOf(c);
      return g && !/^(client details|project details)$/i.test(g) ? g + ' › ' + h : h;
    };
    const rowInfo = row => {
      const v = sheet.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
      const at = names => { const c = mainDealsCol_(cols, names); return c ? String(v[c - 1] || '').trim() : ''; };
      return { project: at(['Project']), unit: at(['Unit Code']), client: at(['Final Client Name', 'Client Name']) };
    };
    const first = Math.max(r0, DATA_START_ROW);
    const info = rowInfo(first);
    if (nR === 1 && nC === 1) {
      const oldV = e.oldValue === undefined ? '' : String(e.oldValue);
      const newV = e.value === undefined ? String(e.range.getDisplayValue() || '') : String(e.value);
      if (oldV === newV) return;
      const name = colName(c0);
      auditLog_(who || 'Sheet user', /^status$/i.test(String(heads[c0 - 1] || '').trim()) ? 'Sheet edit · Status' : 'Sheet edit',
        { project: info.project, unit: info.unit, client: info.client, details: name + ': ' + (oldV || '—') + ' → ' + (newV || '—'), ref: 'Deals row ' + r0 });
    } else {
      const last = r0 + nR - 1;
      auditLog_(who || 'Sheet user', 'Sheet edit (several cells)', {
        project: nR === 1 ? info.project : '', unit: nR === 1 ? info.unit : '', client: nR === 1 ? info.client : '',
        details: (nR * nC) + ' cells changed · ' + colName(c0) + (nC > 1 ? ' … ' + colName(c0 + nC - 1) : ''),
        ref: 'Deals rows ' + first + (last > first ? '–' + last : '') });
    }
  } catch (err) {
    console.warn('auditDealsSheetEdit_: ' + err);
  }
}
