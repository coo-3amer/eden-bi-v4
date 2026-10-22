/************************************************
 * EDEN BI V4 — TelegramBot.js
 *
 * Approve / reject discounts straight from Telegram.
 *  - A deal with a discount above the limit (10%) is posted to the group with
 *    two buttons: ✅ Approve and ⛔ Reject (from the deal form at once, or
 *    within ~5 minutes when it was typed in the sheet).
 *  - Approve = the same as the Discount Approvals page (ticks COO Confirmed).
 *  - Reject: the bot asks for the reason; reply to its message with it.
 *  - /pending in the group lists every discount still waiting, with buttons.
 *  - Only Telegram accounts listed in Discount Approvals › Telegram approvers
 *    (Script Property TELEGRAM_APPROVERS: @usernames or numeric ids) can decide.
 * Telegram is read every minute by the trigger "telegramPollJob", which
 * installs itself from the 5-minute warm-up. Nothing to run by hand.
 ************************************************/

const TG_SENT_KEY_ = 'TG_DISC_SENT';     // deal keys already posted with buttons
const TG_REJ_PREFIX_ = 'TG_REJ_';        // reason prompts waiting for a reply

function tgToken_() { return String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN') || '').trim(); }
function tgChat_() { return String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID') || '').trim(); }
function tgApi_(method, payload) {
  const token = tgToken_();
  if (!token) return null;
  try {
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(payload || {})
    });
    const body = JSON.parse(res.getContentText() || '{}');
    if (!body.ok) console.warn('Telegram ' + method + ': ' + (body.description || res.getResponseCode()));
    return body;
  } catch (e) { console.warn('Telegram ' + method + ': ' + e); return null; }
}

/* ---------- who may decide ---------- */
function tgApprovers_() {
  return String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_APPROVERS') || '')
    .split(/[,\s]+/).map(x => x.trim().replace(/^@/, '').toLowerCase()).filter(Boolean);
}
function tgIsApprover_(from) {
  const list = tgApprovers_();
  if (!from || !list.length) return false;
  return list.indexOf(String(from.id)) >= 0 || (from.username && list.indexOf(String(from.username).toLowerCase()) >= 0);
}
function tgName_(from) { return [from && from.first_name, from && from.last_name].filter(Boolean).join(' ') || (from && from.username) || 'Telegram user'; }
function saveTelegramApprovers(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!discCanApprove_(user)) throw new Error('Only the COO can change who approves from Telegram.');
  const list = String(req && req.approvers || '').split(/[,\s]+/).map(x => x.trim().replace(/^@/, '')).filter(Boolean);
  PropertiesService.getScriptProperties().setProperty('TELEGRAM_APPROVERS', list.join(','));
  auditLog_(user, 'Telegram approvers changed', { details: list.length ? list.map(x => /^\d+$/.test(x) ? x : '@' + x).join(', ') : '(nobody)' });
  return getDiscountApprovals(authToken);
}

/* ---------- posting a discount with buttons ---------- */
function tgSent_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(TG_SENT_KEY_) || '{}'); } catch (e) { return {}; } }
function tgSentSave_(m) {
  const keys = Object.keys(m);
  if (keys.length > 300) keys.sort((a, b) => (m[a].at || 0) - (m[b].at || 0)).slice(0, keys.length - 300).forEach(k => delete m[k]);
  PropertiesService.getScriptProperties().setProperty(TG_SENT_KEY_, JSON.stringify(m));
}
function tgDiscText_(d) {
  const n = v => Math.round(Number(v) || 0).toLocaleString('en-US');
  return ['<b>🏷 Discount waiting for approval — ' + (Math.round(Number(d.pct) * 10) / 10) + '%</b>',
    tgEsc_([d.project, d.unit].filter(Boolean).join(' · ')),
    d.client ? 'Client: ' + tgEsc_(d.client) : '',
    d.total ? 'Price ' + n(d.total) + (d.final ? ' → final ' + n(d.final) : '') : '',
    d.sales ? 'Sales: ' + tgEsc_(d.sales) : ''].filter(Boolean).join('\n');
}
function tgDiscButtons_(row, unit) {
  const u = String(unit || '').slice(0, 40);
  return { inline_keyboard: [[{ text: '✅ Approve', callback_data: ('da|A|' + row + '|' + u).slice(0, 64) },
    { text: '⛔ Reject', callback_data: ('da|R|' + row + '|' + u).slice(0, 64) }]] };
}
/* d = {row, project, unit, client, pct, total, final, sales} */
function tgDiscAnnounce_(d) {
  const chat = tgChat_();
  if (!chat || !tgToken_() || !d || !d.row || !d.unit) return false;
  const key = instKey_(d.project, d.unit), sent = tgSent_();
  if (sent[key]) return false;
  const r = tgApi_('sendMessage', { chat_id: chat, text: tgDiscText_(d), parse_mode: 'HTML', reply_markup: tgDiscButtons_(d.row, d.unit) });
  if (r && r.ok) { sent[key] = { at: Date.now(), msg: r.result.message_id, row: d.row }; tgSentSave_(sent); return true; }
  return false;
}
/* After a decision (page or Telegram): remove the buttons from the posted message. */
function tgDiscMarkDecided_(project, unit, decision, who, reason) {
  const key = instKey_(project, unit), sent = tgSent_(), s = sent[key];
  if (!s || !s.msg || s.done) return;
  tgApi_('editMessageReplyMarkup', { chat_id: tgChat_(), message_id: s.msg, reply_markup: { inline_keyboard: [[{ text: (decision === 'Approved' ? '✅ Approved' : '⛔ Rejected') + ' by ' + String(who).slice(0, 30), callback_data: 'noop' }]] } });
  s.done = decision; tgSentSave_(sent);
}

/* Pending discounts from the dashboard cache (kept fresh by the warm-up). */
function tgPendingDiscounts_() {
  const data = dashCacheRead_('ALL', 60 * 60 * 1000) || buildDashboardData_({ role: 'Admin' });
  const dec = {};
  instRows_(instSheet_(DISC_SHEET_, DISC_HEAD_, false)).forEach(r => { if (r[0]) dec[String(r[0])] = String(r[6] || ''); });
  const T = v => v === true || /^(true|yes|1|✓|✔)$/i.test(String(v || '').trim());
  return (data.rows || []).filter(r => Number(r.discountPct) > DISC_LIMIT_ + 1e-9 && !cancelledStatus_(r.status) && !T(r.cooConfirmed) && !dec[instKey_(r.project, r.unitCode)])
    .map(r => ({ row: Number(r.row) || 0, project: r.project, unit: r.unitCode, client: r.clientName, pct: Number(r.discountPct),
      total: Number(r.primaryTotal) || 0, final: Number(r.amount) || 0, sales: r.salesName, y: Number(r.year) || 0, m: Number(r.month) || 0 }))
    .filter(x => x.row && x.unit);
}
/* From the warm-up: post new pending discounts (deals typed in the sheet). The first run only remembers what is already there. */
function tgDiscWatch_() {
  try {
    if (!tgToken_() || !tgChat_()) return;
    const props = PropertiesService.getScriptProperties();
    const list = tgPendingDiscounts_();
    if (!props.getProperty('TG_DISC_SEEDED')) {
      const sent = tgSent_();
      list.forEach(d => { const k = instKey_(d.project, d.unit); if (!sent[k]) sent[k] = { at: Date.now(), seeded: true }; });
      tgSentSave_(sent); props.setProperty('TG_DISC_SEEDED', '1');
      return;
    }
    list.slice(0, 5).forEach(d => tgDiscAnnounce_(d));
  } catch (e) { console.warn('tgDiscWatch_: ' + e); }
}

/* ---------- reading Telegram (every minute) ---------- */
function telegramPollJob() {
  if (!tgToken_()) return;
  const cache = CacheService.getScriptCache();
  if (cache.get('TG_POLLING')) return;
  cache.put('TG_POLLING', '1', 55);
  try {
    const props = PropertiesService.getScriptProperties();
    const offset = Number(props.getProperty('TG_OFFSET') || 0);
    const r = tgApi_('getUpdates', { offset: offset || undefined, timeout: 0, allowed_updates: ['callback_query', 'message'] });
    if (!r || !r.ok) return;
    let next = offset;
    if (!props.getProperty('TG_OFFSET')) {   // first run: skip anything older than now
      (r.result || []).forEach(u => { next = Math.max(next, u.update_id + 1); });
      props.setProperty('TG_OFFSET', String(next || 1));
      return;
    }
    (r.result || []).forEach(u => {
      next = Math.max(next, u.update_id + 1);
      try {
        if (u.callback_query) tgOnButton_(u.callback_query);
        else if (u.message) tgOnMessage_(u.message);
      } catch (e) { console.warn('Telegram update: ' + e); }
    });
    if (next !== offset) props.setProperty('TG_OFFSET', String(next));
  } finally { cache.remove('TG_POLLING'); }
}

function tgOnButton_(q) {
  const data = String(q.data || '');
  if (data === 'noop') { tgApi_('answerCallbackQuery', { callback_query_id: q.id }); return; }
  if (typeof caOnButton_ === 'function' && caOnButton_(q)) return;
  const m = data.match(/^da\|([AR])\|(\d+)\|(.*)$/);
  if (!m) { tgApi_('answerCallbackQuery', { callback_query_id: q.id }); return; }
  if (!tgIsApprover_(q.from)) {
    tgApi_('answerCallbackQuery', { callback_query_id: q.id, show_alert: true,
      text: 'You are not set as a discount approver. Ask the admin to add ' + (q.from.username ? '@' + q.from.username : q.from.id) + ' in Discount Approvals › Telegram approvers.' });
    return;
  }
  const row = Number(m[2]), unit = m[3];
  const d = tgDealAt_(row, unit);
  if (!d) { tgApi_('answerCallbackQuery', { callback_query_id: q.id, show_alert: true, text: 'This deal moved or changed in the sheet. Use the Discount Approvals page.' }); return; }
  const who = tgName_(q.from) + ' (Telegram)';
  if (m[1] === 'A') {
    try {
      discApply_(who, who, { decision: 'Approved', dealRow: row, unitCode: d.unit, project: d.project, client: d.client, discount: d.pct });
      tgApi_('answerCallbackQuery', { callback_query_id: q.id, text: '✅ Approved' });
      if (q.message) tgApi_('editMessageReplyMarkup', { chat_id: q.message.chat.id, message_id: q.message.message_id, reply_markup: { inline_keyboard: [[{ text: '✅ Approved by ' + tgName_(q.from).slice(0, 30), callback_data: 'noop' }]] } });
    } catch (e) { tgApi_('answerCallbackQuery', { callback_query_id: q.id, show_alert: true, text: String(e && e.message || e).slice(0, 190) }); }
    return;
  }
  // Reject: ask for the reason
  tgApi_('answerCallbackQuery', { callback_query_id: q.id, text: 'Write the reason' });
  const ask = tgApi_('sendMessage', { chat_id: q.message ? q.message.chat.id : tgChat_(),
    text: '⛔ Reason for rejecting the discount on ' + d.unit + ' (' + (d.client || '') + ')?\nReply to this message with the reason.',
    reply_markup: { force_reply: true, selective: true }, reply_to_message_id: q.message ? q.message.message_id : undefined });
  if (ask && ask.ok) PropertiesService.getScriptProperties().setProperty(TG_REJ_PREFIX_ + ask.result.message_id,
    JSON.stringify({ row: row, unit: d.unit, by: q.from.id, msg: q.message ? q.message.message_id : 0, chat: q.message ? q.message.chat.id : '' }));
}

function tgOnMessage_(msg) {
  const text = String(msg.text || '').trim();
  const props = PropertiesService.getScriptProperties();
  // reply with a rejection reason
  if (msg.reply_to_message) {
    const k = TG_REJ_PREFIX_ + msg.reply_to_message.message_id, raw = props.getProperty(k);
    if (raw) {
      const p = JSON.parse(raw);
      if (!tgIsApprover_(msg.from)) return;
      if (!text) return;
      const d = tgDealAt_(p.row, p.unit);
      if (!d) { tgApi_('sendMessage', { chat_id: msg.chat.id, text: 'This deal moved or changed in the sheet. Use the Discount Approvals page.', reply_to_message_id: msg.message_id }); return; }
      const who = tgName_(msg.from) + ' (Telegram)';
      try {
        discApply_(who, who, { decision: 'Rejected', reason: text, dealRow: p.row, unitCode: d.unit, project: d.project, client: d.client, discount: d.pct });
        props.deleteProperty(k);
        if (p.msg) tgApi_('editMessageReplyMarkup', { chat_id: p.chat || msg.chat.id, message_id: p.msg, reply_markup: { inline_keyboard: [[{ text: '⛔ Rejected by ' + tgName_(msg.from).slice(0, 30), callback_data: 'noop' }]] } });
        tgApi_('sendMessage', { chat_id: msg.chat.id, text: '⛔ Rejected — ' + d.unit + '. Reason saved.', reply_to_message_id: msg.message_id });
      } catch (e) { tgApi_('sendMessage', { chat_id: msg.chat.id, text: String(e && e.message || e).slice(0, 300), reply_to_message_id: msg.message_id }); }
      return;
    }
  }
  const cmd = (text.match(/^\/(\w+)/) || [])[1];
  if (!cmd) return;
  if (cmd === 'myid') {
    tgApi_('sendMessage', { chat_id: msg.chat.id, reply_to_message_id: msg.message_id,
      text: 'Your Telegram id: ' + msg.from.id + (msg.from.username ? ' (@' + msg.from.username + ')' : '') + (tgIsApprover_(msg.from) ? '\n✅ You can approve discounts.' : '\nNot a discount approver yet.') });
  } else if (cmd === 'pending') {
    const list = tgPendingDiscounts_();
    if (!list.length) { tgApi_('sendMessage', { chat_id: msg.chat.id, text: '✅ No discounts waiting for approval.' }); return; }
    tgApi_('sendMessage', { chat_id: msg.chat.id, text: list.length + ' discount' + (list.length > 1 ? 's' : '') + ' waiting for approval' + (list.length > 10 ? ' — showing the newest 10' : '') + ':' });
    const sent = tgSent_();
    list.sort((a, b) => (b.y * 100 + b.m) - (a.y * 100 + a.m) || b.row - a.row).slice(0, 10).forEach(d => {
      const r = tgApi_('sendMessage', { chat_id: msg.chat.id, text: tgDiscText_(d), parse_mode: 'HTML', reply_markup: tgDiscButtons_(d.row, d.unit) });
      if (r && r.ok) sent[instKey_(d.project, d.unit)] = { at: Date.now(), msg: r.result.message_id, row: d.row };
    });
    tgSentSave_(sent);
  } else if (cmd === 'changes') {
    const sh = caSheet_(), last = sh.getLastRow();
    const pend = last < 2 ? [] : sh.getRange(2, 1, last - 1, CA_HEAD_.length).getValues().filter(r => String(r[11]) === 'Pending');
    if (!pend.length) { tgApi_('sendMessage', { chat_id: msg.chat.id, text: '✅ No sheet changes waiting for approval.' }); return; }
    pend.slice(-10).forEach(r => tgApi_('sendMessage', { chat_id: msg.chat.id, text: caText_(r), parse_mode: 'HTML',
      reply_markup: { inline_keyboard: [[{ text: '✅ Approve', callback_data: 'ca|A|' + r[0] }, { text: '⛔ Reject & undo', callback_data: 'ca|R|' + r[0] }]] } }));
  } else if (cmd === 'help' || cmd === 'start') {
    tgApi_('sendMessage', { chat_id: msg.chat.id, text: 'EDEN BI bot\n/pending — discounts waiting for approval (with buttons)\n/changes — sheet changes waiting for approval\n/myid — your Telegram id' });
  }
}

/* Reads the deal on that Deals row and checks it is still the same unit. */
function tgDealAt_(row, unit) {
  const sheet = getMainDealsSheet_();
  const cols = mainDealsColumns_(sheet);
  if (!row || row < DATA_START_ROW || row > sheet.getLastRow()) return null;
  const v = sheet.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
  const at = names => { const c = mainDealsCol_(cols, names); return c ? String(v[c - 1] || '').trim() : ''; };
  const live = at(['Unit Code']).toUpperCase().replace(/\.+$/, '');
  if (!live || live !== String(unit || '').trim().toUpperCase().replace(/\.+$/, '')) return null;
  const pend = tgPendingDiscounts_().find(d => d.row === row);
  return { unit: at(['Unit Code']), project: at(['Project']), client: at(['Final Client Name', 'Client Name']), pct: pend ? pend.pct : '' };
}

/* From the warm-up: make sure the every-minute reader exists (checked hourly). */
function ensureTelegramPollTrigger_() {
  try {
    if (!tgToken_()) return;
    const cache = CacheService.getScriptCache();
    if (cache.get('TG_TRIGGER_OK')) return;
    const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'telegramPollJob');
    if (!has) ScriptApp.newTrigger('telegramPollJob').timeBased().everyMinutes(1).create();
    cache.put('TG_TRIGGER_OK', '1', 3600);
  } catch (e) { console.warn('ensureTelegramPollTrigger_: ' + e); }
}
