/************************************************
 * EDEN BI V4 — Notify.js
 *
 * Telegram alerts to the management group.
 * One-time setup (owner):
 *   1. Script Properties: TELEGRAM_BOT_TOKEN = the token from @BotFather.
 *   2. Add the bot to the group and send any message there.
 *   3. In the Apps Script editor choose "setupTelegram" and press ▶ Run,
 *      then press Allow. It finds the group, saves TELEGRAM_CHAT_ID and
 *      sends a test message.
 * Alerts come from the Activity Log (auditLog_), so every action that is
 * logged can be announced. Sending never blocks or breaks a save.
 ************************************************/

const TG_EVENTS_ = {
  'Deal added': '🆕',
  'Deal status changed': '🔄',
  'Deal status changed (sheet)': '🔄',
  'Deal added (sheet)': '🆕',
  'Inventory edit (sheet)': '🏢',
  'Payment recorded': '💰',
  'Payment deleted': '🗑',
  'Discount approved': '✅',
  'Discount rejected': '⛔',
  'Broker commission paid': '🤝',
  'Prices changed': '🏷',
  'Daily backup': '⚠️'      // only when something failed
};

function tgEsc_(v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/* Sends one message. Returns true when Telegram accepted it. Never throws. */
function notifyTelegram_(html) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = String(props.getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
    const chat = String(props.getProperty('TELEGRAM_CHAT_ID') || '').trim();
    if (!token || !chat || props.getProperty('TELEGRAM_OFF') === '1') return false;
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: chat, text: String(html).slice(0, 3900), parse_mode: 'HTML', disable_web_page_preview: true })
    });
    if (res.getResponseCode() !== 200) console.warn('Telegram: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
    return res.getResponseCode() === 200;
  } catch (e) {
    console.warn('notifyTelegram_: ' + e);
    return false;
  }
}

/* Called by auditLog_ for every logged action; announces the important ones. */
function notifyFromAudit_(who, action, info) {
  try {
    info = info || {};
    const icon = TG_EVENTS_[action];
    if (!icon) return;
    const details = String(info.details || '');
    if (action === 'Daily backup' && details.indexOf('Failed') < 0) return;
    const head = [info.project, info.unit].filter(Boolean).join(' · ');
    const lines = ['<b>' + icon + ' ' + tgEsc_(action) + '</b>' + (head ? ' — ' + tgEsc_(head) : '')];
    if (info.client) lines.push('Client: ' + tgEsc_(info.client));
    if (details) lines.push(tgEsc_(details.length > 600 ? details.slice(0, 600) + '…' : details));
    if (Number(info.discount) > (typeof DISC_LIMIT_ !== 'undefined' ? DISC_LIMIT_ : 10)) {
      lines.push('⚠️ <b>Discount ' + (Math.round(Number(info.discount) * 10) / 10) + '%</b> — waiting for COO approval (Transactions › Discount Approvals)');
    }
    if (who) lines.push('<i>by ' + tgEsc_(who) + '</i>');
    notifyTelegram_(lines.join('\n'));
  } catch (e) {
    console.warn('notifyFromAudit_: ' + e);
  }
}

/* Run once from the editor (▶ Run): finds the group, saves its id, sends a test message. */
function setupTelegram() {
  const props = PropertiesService.getScriptProperties();
  const token = String(props.getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  if (!token) throw new Error('Add the Script Property TELEGRAM_BOT_TOKEN first (Project Settings › Script Properties).');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates', { muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText() || '{}');
  if (!body.ok) throw new Error('Telegram did not accept the token: ' + (body.description || res.getResponseCode()) + '. Check TELEGRAM_BOT_TOKEN.');
  const chats = {};
  (body.result || []).forEach(u => {
    const m = u.message || u.channel_post || u.my_chat_member || u.edited_message;
    const c = m && m.chat;
    if (c && c.id) chats[c.id] = { id: c.id, type: c.type, title: c.title || [c.first_name, c.last_name].filter(Boolean).join(' ') };
  });
  const list = Object.keys(chats).map(k => chats[k]);
  const group = list.filter(c => c.type === 'group' || c.type === 'supergroup').pop() || list.pop();
  if (!group) {
    if (props.getProperty('TELEGRAM_CHAT_ID')) {
      const ok = notifyTelegram_('✅ <b>EDEN BI</b> — test message. Alerts are on.');
      console.log(ok ? 'Test message sent to the saved group.' : 'Could not send to the saved group.');
      return ok ? 'Test message sent.' : 'Could not send — check the bot is still in the group.';
    }
    throw new Error('No group found. Add the bot to the group, send any message in the group, then run setupTelegram again.');
  }
  props.setProperty('TELEGRAM_CHAT_ID', String(group.id));
  const ok = notifyTelegram_('✅ <b>EDEN BI is connected</b>\nThis group will now get alerts: new deals, status changes and cancellations, payments, discounts above 10%, broker payouts and price changes.');
  const msg = (ok ? 'Connected to "' : 'Saved but the test message failed for "') + group.title + '" (' + group.id + ').';
  console.log(msg);
  return msg;
}
