/************************************************
 * EDEN BI V4 — DiscountApprovalData.js
 *
 * Deals with a discount above the limit (10%) wait for the COO.
 *  - Approve: ticks "COO Confirmed" (Deal Confirmation group) on the deal's
 *    row in the Deals sheet, and records the decision.
 *  - Reject: records the decision and the reason, and writes a note on the
 *    deal's Discount cell. The deal itself is not changed.
 * Every decision goes to the "Discount Approvals" tab and the Activity Log.
 * Who can decide: admins. To limit it to named people, set the Script
 * Property DISCOUNT_APPROVERS = usernames separated by commas.
 ************************************************/

const DISC_LIMIT_ = 10;   // % — a discount ABOVE this needs approval
const DISC_SHEET_ = 'Discount Approvals';
const DISC_HEAD_ = ['Deal Key', 'Deal Row', 'Project', 'Unit Code', 'Client', 'Discount %', 'Decision', 'Reason', 'Decided By', 'Decided At'];

function discCanApprove_(user) {
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') return false;
  const list = String(PropertiesService.getScriptProperties().getProperty('DISCOUNT_APPROVERS') || '').trim();
  if (!list) return true;
  const me = [user.username, user.name, user.email].map(x => String(x || '').trim().toLowerCase()).filter(Boolean);
  return list.split(',').map(x => x.trim().toLowerCase()).filter(Boolean).some(x => me.indexOf(x) >= 0);
}

/* Decisions taken so far (latest per deal) + whether this user may decide. */
function getDiscountApprovals(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Discount approvals are for admins only.');
  const decisions = {};
  instRows_(instSheet_(DISC_SHEET_, DISC_HEAD_, false)).forEach(r => {
    const k = String(r[0] || '').trim();
    if (!k) return;
    decisions[k] = { dealRow: Number(r[1]) || 0, project: String(r[2] || ''), unitCode: String(r[3] || ''), client: String(r[4] || ''),
      discount: Number(r[5]) || 0, decision: String(r[6] || ''), reason: String(r[7] || ''), by: String(r[8] || ''), at: instDateOut_(r[9]) };
  });
  return { limit: DISC_LIMIT_, canApprove: discCanApprove_(user), decisions: decisions };
}

/* decision: 'Approved' | 'Rejected'. */
function decideDiscount(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!discCanApprove_(user)) throw new Error('Only the COO can approve or reject discounts.');
  req = req || {};
  const decision = /^approve/i.test(String(req.decision || '')) ? 'Approved' : /^reject/i.test(String(req.decision || '')) ? 'Rejected' : '';
  if (!decision) throw new Error('Choose Approve or Reject.');
  const reason = String(req.reason || '').trim();
  if (decision === 'Rejected' && !reason) throw new Error('Write the reason for rejecting.');
  const row = Number(req.dealRow);
  if (!row || row < DATA_START_ROW || !req.unitCode) throw new Error('Deal is missing.');
  const who = String(user.name || user.username || 'COO');
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  let msg = '';

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    // The row must still be this deal.
    const sheet = getMainDealsSheet_();
    const cols = mainDealsColumns_(sheet);
    const unitCol = mainDealsCol_(cols, ['Unit Code']);
    const liveUnit = unitCol ? String(sheet.getRange(row, unitCol).getDisplayValue()).trim().toUpperCase().replace(/\.+$/, '') : '';
    if (liveUnit !== String(req.unitCode).trim().toUpperCase().replace(/\.+$/, '')) throw new Error('The Deals sheet changed (rows moved). Refresh and try again.');

    if (decision === 'Approved') {
      const cooCol = mainDealsCol_(cols, ['deal confirmation|coo confirmed', 'COO Confirmed']);
      if (cooCol) { sheet.getRange(row, cooCol).setValue(true); msg = 'COO Confirmed ticked.'; }
      else msg = 'Approved (the "COO Confirmed" column was not found, so nothing was ticked).';
    }
    const discCol = mainDealsCol_(cols, ['price details|discount', 'Discount']);
    if (discCol) {
      try {
        const cell = sheet.getRange(row, discCol), prev = cell.getNote();
        cell.setNote((prev ? prev + '\n' : '') + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm') + ' • Discount ' + decision.toLowerCase() + ' by ' + who + (reason ? ': ' + reason : ''));
      } catch (e) {}
    }

    const sh = instSheet_(DISC_SHEET_, DISC_HEAD_, true);
    const key = instKey_(req.project, req.unitCode);
    const line = [key, row, req.project, req.unitCode, String(req.client || ''), Number(req.discount) || '', decision, reason, who, new Date()];
    const rows = instRows_(sh);
    const i = rows.findIndex(r => String(r[0]) === key);
    if (i >= 0) sh.getRange(i + 2, 1, 1, line.length).setValues([line]);
    else sh.appendRow(line);
  } finally { lock.releaseLock(); }

  auditLog_(user, 'Discount ' + decision.toLowerCase(), { project: req.project, unit: req.unitCode, client: req.client,
    details: 'Discount ' + (Number(req.discount) || 0) + '%' + (reason ? ' · ' + reason : ''), ref: 'Deals row ' + row });
  const out = getDiscountApprovals(authToken);
  out.message = msg;
  return out;
}
