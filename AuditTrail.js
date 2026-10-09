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
