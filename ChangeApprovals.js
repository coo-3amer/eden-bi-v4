/************************************************
 * EDEN BI V4 — ChangeApprovals.js
 *
 * Sensitive hand edits in the Deals sheet wait for approval:
 *   always   : Cancel Details › Date / Reason / Received money (cancellation, refund)
 *   after the contract (Contract Details › Date filled):
 *              Final Price, Discount, Meter Price After Discount, Unit Code,
 *              Sales Name, Shared With
 * The edit stays in the sheet; a request is logged in the hidden sheet
 * "_Change Approvals" and posted on Telegram with ✅ / ⛔ buttons.
 *   Approve → kept.   Reject → the cell goes back to its old value.
 * Edits by people in CHANGE_APPROVAL_EXEMPT (emails, comma separated —
 * default coo@edenalkhalij.com) don't need approval. CHANGE_APPROVAL_OFF=1
 * turns the whole thing off. Approvers = the Telegram discount approvers
 * and the admins on the dashboard.
 ************************************************/

const CA_SHEET_ = '_Change Approvals';
const CA_HEAD_ = ['ID', 'Deals Row', 'Project', 'Unit', 'Client', 'Field', 'Column', 'Old', 'New', 'By', 'At', 'Status', 'Decided By', 'Decided At', 'Note'];
const CA_RULES_ = [
  { names: ['cancel details|date'], label: 'Cancellation date', always: true },
  { names: ['cancel details|reason'], label: 'Cancellation reason', always: true },
  { names: ['cancel details|received money'], label: 'Refund (received money)', always: true },
  { names: ['price details|final price', 'Final Price'], label: 'Final price' },
  { names: ['price details|discount', 'Discount'], label: 'Discount' },
  { names: ['price details|meter price after discount', 'Meter Price After Discount'], label: 'Meter price after discount' },
  { names: ['Unit Code'], label: 'Unit code' },
  { names: ['transaction details|sales name', 'Sales Name'], label: 'Sales name' },
  { names: ['transaction details|shared with', 'Shared With'], label: 'Shared with' }
];

function caSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(CA_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(CA_SHEET_);
    sh.getRange(1, 1, 1, CA_HEAD_.length).setValues([CA_HEAD_]).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}
/* Column numbers of the watched fields + contract date (cached 10 min). */
function caCols_(sheet) {
  const cache = CacheService.getScriptCache(), hit = cache.get('CA_COLS_V1');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const cols = mainDealsColumns_(sheet), at = n => mainDealsCol_(cols, n);
  const out = { contract: at(['contract details|date']), project: at(['Project']), unit: at(['Unit Code']), client: at(['Final Client Name', 'Client Name']), watch: {} };
  CA_RULES_.forEach((r, i) => { const c = at(r.names); if (c && !out.watch[c]) out.watch[c] = i; });
  cache.put('CA_COLS_V1', JSON.stringify(out), 600);
  return out;
}
function caExempt_(email) {
  const raw = PropertiesService.getScriptProperties().getProperty('CHANGE_APPROVAL_EXEMPT');
  const list = String(raw == null ? 'coo@edenalkhalij.com' : raw).split(/[,\s]+/).map(x => x.trim().toLowerCase()).filter(Boolean);
  return !!email && list.indexOf(String(email).toLowerCase()) >= 0;
}

/* From the Deals onEdit (LiveSync). Single-cell edits only (old value known). */
function caWatchEdit_(e) {
  try {
    if (!e || !e.range || PropertiesService.getScriptProperties().getProperty('CHANGE_APPROVAL_OFF') === '1') return;
    if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
    const sheet = e.range.getSheet(), row = e.range.getRow(), col = e.range.getColumn();
    if (row < DATA_START_ROW) return;
    const C = caCols_(sheet), idx = C.watch[col];
    if (idx === undefined) return;
    const rule = CA_RULES_[idx];
    const oldV = e.oldValue === undefined ? '' : String(e.oldValue), newV = String(e.range.getDisplayValue() || '');
    if (oldV === newV || caShow_(rule.label, oldV) === newV) return;
    let who = '';
    try { who = (e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (err) {}
    if (caExempt_(who)) return;
    const vals = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const get = c => c ? String(vals[c - 1] || '').trim() : '';
    if (!rule.always && !get(C.contract)) return;                 // price/sales edits before the contract are free
    const id = 'C' + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 4).toUpperCase();
    const rec = [id, row, get(C.project), get(C.unit), get(C.client), rule.label, col, oldV, newV, who || 'Sheet user', new Date(), 'Pending', '', '', ''];
    const sh = caSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, 1, CA_HEAD_.length).setValues([rec]);
    caAnnounce_(rec);
  } catch (err) { console.warn('caWatchEdit_: ' + err); }
}
/* Old values come raw from onEdit: a date is a serial number, a checkbox TRUE/FALSE. */
function caShow_(label, v) {
  const t = String(v == null ? '' : v);
  if (/date/i.test(label) && /^\d{4,6}(\.\d+)?$/.test(t)) return Utilities.formatDate(new Date(Math.round((Number(t) - 25569) * 864e5)), 'UTC', 'd-MMM-yyyy');
  return t;
}
function caRaw_(v) {
  const t = String(v == null ? '' : v);
  if (t === 'TRUE' || t === 'FALSE') return t === 'TRUE';
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  return t;
}
function caText_(r) {
  const E = typeof tgEsc_ === 'function' ? tgEsc_ : (v => String(v));
  return ['<b>🔐 Change waiting for approval — ' + E(r[5]) + '</b>', E([r[2], r[3]].filter(Boolean).join(' · ')), r[4] ? 'Client: ' + E(r[4]) : '',
    E(caShow_(r[5], r[7]) || '—') + '  →  <b>' + E(r[8] || '—') + '</b>', 'By ' + E(r[9]) + ' · Deals row ' + r[1]].filter(Boolean).join('\n');
}
function caAnnounce_(r) {
  const chat = typeof tgChat_ === 'function' ? tgChat_() : '';
  if (!chat || typeof tgApi_ !== 'function' || PropertiesService.getScriptProperties().getProperty('TELEGRAM_OFF') === '1') return;
  tgApi_('sendMessage', { chat_id: chat, text: caText_(r), parse_mode: 'HTML', disable_web_page_preview: true,
    reply_markup: { inline_keyboard: [[{ text: '✅ Approve', callback_data: 'ca|A|' + r[0] }, { text: '⛔ Reject & undo', callback_data: 'ca|R|' + r[0] }]] } });
}

/* Approve / reject. Reject puts the old value back if the cell still holds the new one. */
function caDecide_(id, approve, who, note) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = caSheet_(), last = sh.getLastRow();
    if (last < 2) throw new Error('Request not found.');
    const rows = sh.getRange(2, 1, last - 1, CA_HEAD_.length).getValues();
    const i = rows.findIndex(r => String(r[0]) === String(id));
    if (i < 0) throw new Error('Request not found.');
    const r = rows[i];
    if (String(r[11]) !== 'Pending') throw new Error('Already ' + String(r[11]).toLowerCase() + (r[12] ? ' by ' + r[12] : '') + '.');
    let status = approve ? 'Approved' : 'Rejected', extra = '';
    if (!approve) {
      const deals = getMainDealsSheet_(), cell = deals.getRange(Number(r[1]), Number(r[6]));
      const C = caCols_(deals), unitNow = String(deals.getRange(Number(r[1]), C.unit).getDisplayValue() || '').trim();
      const sameUnit = r[5] === 'Unit code' ? unitNow === String(r[8]).trim() : unitNow === String(r[3]).trim();
      if (!sameUnit) throw new Error('The row changed in the sheet (unit is now ' + (unitNow || 'empty') + '). Undo it by hand.');
      const now = String(cell.getDisplayValue() || '');
      if (now !== String(r[8])) { status = 'Rejected'; extra = 'cell changed again since — not undone'; }
      else { cell.setValue(r[7] === '' ? '' : caRaw_(r[7])); extra = 'undone'; try { clearDashboardCache_(); CacheService.getScriptCache().put('LIVE_DIRTY_DASH', String(Date.now()), 600); } catch (e) {} }
    }
    sh.getRange(i + 2, 12, 1, 4).setValues([[status, who, new Date(), [note, extra].filter(Boolean).join(' · ')]]);
    auditLog_(who, 'Change ' + status.toLowerCase(), { project: String(r[2]), unit: String(r[3]), client: String(r[4]),
      details: r[5] + ': ' + (caShow_(r[5], r[7]) || '—') + ' → ' + (r[8] || '—') + (extra ? ' · ' + extra : '') + (note ? ' · ' + note : ''), ref: 'Deals row ' + r[1] });
    return { id: id, status: status, extra: extra };
  } finally { lock.releaseLock(); }
}

/* Telegram button (called from tgOnButton_). Returns true when handled. */
function caOnButton_(q) {
  const m = String(q.data || '').match(/^ca\|([AR])\|(\w+)$/);
  if (!m) return false;
  if (!tgIsApprover_(q.from)) {
    tgApi_('answerCallbackQuery', { callback_query_id: q.id, show_alert: true, text: 'You are not an approver. Ask the admin to add ' + (q.from.username ? '@' + q.from.username : q.from.id) + ' in Discount Approvals › Telegram approvers.' });
    return true;
  }
  try {
    const res = caDecide_(m[2], m[1] === 'A', tgName_(q.from) + ' (Telegram)', '');
    const label = (res.status === 'Approved' ? '✅ Approved' : '⛔ Rejected' + (res.extra === 'undone' ? ' & undone' : '')) + ' by ' + tgName_(q.from).slice(0, 25);
    tgApi_('answerCallbackQuery', { callback_query_id: q.id, text: label });
    if (q.message) tgApi_('editMessageReplyMarkup', { chat_id: q.message.chat.id, message_id: q.message.message_id, reply_markup: { inline_keyboard: [[{ text: label, callback_data: 'noop' }]] } });
  } catch (e) { tgApi_('answerCallbackQuery', { callback_query_id: q.id, show_alert: true, text: String(e && e.message || e).slice(0, 190) }); }
  return true;
}

/* Dashboard (admins). */
function getChangeApprovals(authToken) {
  const user = targetsRequireAdmin_(authToken);
  const sh = caSheet_(), last = sh.getLastRow();
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, CA_HEAD_.length).getValues().map(r => ({
    id: String(r[0]), row: Number(r[1]), project: String(r[2]), unit: String(r[3]), client: String(r[4]), field: String(r[5]),
    old: caShow_(String(r[5]), r[7]), now: String(r[8]), by: String(r[9]), at: r[10] instanceof Date ? r[10].toISOString() : String(r[10]),
    status: String(r[11]), decidedBy: String(r[12]), decidedAt: r[13] instanceof Date ? r[13].toISOString() : String(r[13] || ''), note: String(r[14] || '') }));
  const props = PropertiesService.getScriptProperties();
  return { rows: rows.reverse().slice(0, 300), off: props.getProperty('CHANGE_APPROVAL_OFF') === '1',
    exempt: String(props.getProperty('CHANGE_APPROVAL_EXEMPT') == null ? 'coo@edenalkhalij.com' : props.getProperty('CHANGE_APPROVAL_EXEMPT')), me: user.name || user.username };
}
function decideChange(authToken, req) {
  const user = targetsRequireAdmin_(authToken);
  req = req || {};
  caDecide_(req.id, !!req.approve, user.name || user.username, String(req.note || '').slice(0, 300));
  return getChangeApprovals(authToken);
}
function saveChangeApprovalSettings(authToken, req) {
  const user = targetsRequireAdmin_(authToken);
  const props = PropertiesService.getScriptProperties();
  req = req || {};
  if (req.off !== undefined) req.off ? props.setProperty('CHANGE_APPROVAL_OFF', '1') : props.deleteProperty('CHANGE_APPROVAL_OFF');
  if (req.exempt !== undefined) props.setProperty('CHANGE_APPROVAL_EXEMPT', String(req.exempt || '').trim());
  auditLog_(user, 'Change approvals settings', { details: (req.off ? 'off' : 'on') + ' · exempt: ' + String(req.exempt || '') });
  return getChangeApprovals(authToken);
}
