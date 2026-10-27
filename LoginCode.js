/************************************************
 * EDEN BI V4 — LoginCode.js
 *
 * Second step at sign-in: after the password, a 6-digit code is sent to
 * the user's Telegram (EDEN BI bot). First time, the user links Telegram
 * once by opening the bot from the login screen.
 *  - Who must use it: Script Property LOGIN_2FA = off | admins | all
 *    (not set = admins). Changed from System › Users.
 *  - "Remember this device" skips the code for 30 days on that browser.
 *  - Changing a user's password forgets their remembered devices.
 *  - If the bot token is missing the code step is skipped (no lock-out).
 ************************************************/

const LC_CODE_MIN_ = 5;          // code is valid for 5 minutes
const LC_MAX_TRIES_ = 5;
const LC_DEVICE_DAYS_ = 30;

function lcMode_() {
  const v = String(PropertiesService.getScriptProperties().getProperty('LOGIN_2FA') || 'admins').trim().toLowerCase();
  return ['off', 'admins', 'all'].indexOf(v) >= 0 ? v : 'admins';
}
function lcNeeded_(user) {
  const mode = lcMode_();
  if (mode === 'off' || typeof tgToken_ !== 'function' || !tgToken_()) return false;
  if (mode === 'all') return true;
  const role = String(user && user.role || '').trim().toLowerCase();
  return role === 'admin' || role === 'super admin';
}

/* ---------- Users sheet: "Telegram ID" column ---------- */
function lcUserCol_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const head = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  let t = head.indexOf('Telegram ID') + 1;
  if (!t) { t = lastCol + 1; sheet.getRange(1, t).setValue('Telegram ID'); }
  return { u: head.indexOf('Username') + 1, t: t };
}
function lcFindUserRow_(sheet, cols, username) {
  const last = sheet.getLastRow();
  if (last < 2 || !cols.u) return 0;
  const names = sheet.getRange(2, cols.u, last - 1, 1).getDisplayValues();
  for (let i = 0; i < names.length; i++) if (String(names[i][0]).trim().toLowerCase() === username) return i + 2;
  return 0;
}
function lcGetChat_(username) {
  const sheet = getUsersSheet_(), cols = lcUserCol_(sheet), row = lcFindUserRow_(sheet, cols, username);
  return row ? String(sheet.getRange(row, cols.t).getDisplayValue() || '').trim() : '';
}
function lcSetChat_(username, chatId) {
  const sheet = getUsersSheet_(), cols = lcUserCol_(sheet), row = lcFindUserRow_(sheet, cols, username);
  if (!row) return false;
  sheet.getRange(row, cols.t).setNumberFormat('@').setValue(String(chatId || ''));
  return true;
}

/* ---------- trusted devices ---------- */
function lcDevKey_(t) { return 'LCDEV_' + String(t || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 64); }
function lcDeviceOk_(username, deviceToken) {
  if (!deviceToken) return false;
  const raw = PropertiesService.getScriptProperties().getProperty(lcDevKey_(deviceToken));
  if (!raw) return false;
  try { const d = JSON.parse(raw); return d.u === username && Number(d.exp) > Date.now(); } catch (e) { return false; }
}
function lcRemember_(username) {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const props = PropertiesService.getScriptProperties();
  props.setProperty(lcDevKey_(t), JSON.stringify({ u: username, exp: Date.now() + LC_DEVICE_DAYS_ * 864e5 }));
  // tidy expired ones
  try {
    const all = props.getProperties();
    Object.keys(all).forEach(k => { if (k.indexOf('LCDEV_') === 0) { try { if (Number(JSON.parse(all[k]).exp) < Date.now()) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } } });
  } catch (e) {}
  return t;
}
function lcForgetDevices_(username) {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties();
  Object.keys(all).forEach(k => { if (k.indexOf('LCDEV_') === 0) { try { if (JSON.parse(all[k]).u === username) props.deleteProperty(k); } catch (e) {} } });
}

/* ---------- pending sign-ins (cache, 10 minutes) ---------- */
function lcGet_(pid) { const v = CacheService.getScriptCache().get('LCP_' + String(pid || '')); return v ? JSON.parse(v) : null; }
function lcPut_(pid, p) { CacheService.getScriptCache().put('LCP_' + pid, JSON.stringify(p), 600); }
function lcDrop_(pid) { CacheService.getScriptCache().remove('LCP_' + pid); }

function lcBotName_() {
  const props = PropertiesService.getScriptProperties();
  let n = props.getProperty('TG_BOT_USERNAME');
  if (!n) { const r = tgApi_('getMe', {}); n = r && r.ok && r.result ? r.result.username : ''; if (n) props.setProperty('TG_BOT_USERNAME', n); }
  return n || '';
}
function lcSendCode_(p) {
  p.code = String(Math.floor(100000 + Math.random() * 900000));
  p.exp = Date.now() + LC_CODE_MIN_ * 60000;
  p.tries = 0;
  p.sentAt = Date.now();
  const when = Utilities.formatDate(new Date(), 'Africa/Cairo', 'HH:mm');
  const r = tgApi_('sendMessage', { chat_id: p.chat, parse_mode: 'HTML',
    text: '🔐 <b>كود الدخول لـ EDEN BI</b>\n\n<code>' + p.code + '</code>\n\nللحساب: ' + p.user.username + ' · الساعة ' + when +
      '\nصالح ' + LC_CODE_MIN_ + ' دقايق.\n\nلو مش إنت اللي بتحاول تدخل، متديش الكود لحد وغيّر الباسورد فوراً.' });
  return !!(r && r.ok);
}
function lcMask_(chat) { const s = String(chat || ''); return s.length > 4 ? '•••' + s.slice(-3) : 'Telegram'; }

/* Called by loginUser after the password is right. Returns a step, or null to sign in now. */
function lcStart_(user, deviceToken) {
  if (!lcNeeded_(user)) return null;
  if (lcDeviceOk_(user.username, deviceToken)) return null;
  const pid = Utilities.getUuid().replace(/-/g, '');
  const chat = lcGetChat_(user.username);
  const p = { user: user, chat: chat };
  if (chat) {
    if (!lcSendCode_(p)) {   // chat blocked the bot / was deleted → link again
      p.chat = '';
    } else {
      lcPut_(pid, p);
      return { success: false, step: 'code', pid: pid, to: lcMask_(chat) };
    }
  }
  p.link = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  lcPut_(pid, p);
  CacheService.getScriptCache().put('LCL_' + p.link, pid, 600);
  const bot = lcBotName_();
  return { success: false, step: 'link', pid: pid, bot: bot, link: bot ? 'https://t.me/' + bot + '?start=L' + p.link : '' };
}

/* Telegram: /start L<link> from the login screen button. */
function lcOnTelegramLink_(msg, payload) {
  const cache = CacheService.getScriptCache();
  const pid = cache.get('LCL_' + String(payload || '').replace(/^L/, ''));
  const p = pid ? lcGet_(pid) : null;
  if (!p) { tgApi_('sendMessage', { chat_id: msg.chat.id, text: 'الرابط ده انتهى. ارجع لصفحة الدخول وسجّل تاني.' }); return; }
  if (msg.chat.type !== 'private') { tgApi_('sendMessage', { chat_id: msg.chat.id, text: 'افتح البوت في شات خاص مش جروب.' }); return; }
  p.chat = String(msg.chat.id);
  lcSetChat_(p.user.username, p.chat);
  tgApi_('sendMessage', { chat_id: p.chat, text: '✅ تم ربط تيليجرام بحساب ' + p.user.username + ' في EDEN BI.\nمن هنا ورايح كود الدخول هيوصلك هنا.' });
  lcSendCode_(p);
  lcPut_(pid, p);
  cache.remove('LCL_' + String(payload).replace(/^L/, ''));
  if (typeof auditLog_ === 'function') auditLog_(p.user, 'Telegram linked for sign-in', { details: tgName_(msg.from) });
}

/* ---------- called from the login screen (no session yet) ---------- */
function checkLoginLink(pid) {
  let p = lcGet_(pid);
  if (!p) return { step: 'expired' };
  if (!p.chat) {   // read Telegram now instead of waiting for the 1-minute job
    try { telegramPollJob(); } catch (e) {}
    p = lcGet_(pid);
    if (!p) return { step: 'expired' };
  }
  return p.chat && p.code ? { step: 'code', to: lcMask_(p.chat) } : { step: 'link' };
}
function resendLoginCode(pid) {
  const p = lcGet_(pid);
  if (!p || !p.chat) return { ok: false, message: 'Please sign in again.' };
  if (Date.now() - Number(p.sentAt || 0) < 30000) return { ok: false, message: 'Wait a few seconds before asking for a new code.' };
  const ok = lcSendCode_(p);
  lcPut_(pid, p);
  return { ok: ok, message: ok ? 'A new code was sent.' : 'Telegram did not accept the message.' };
}
function verifyLoginCode(pid, code, remember) {
  const p = lcGet_(pid);
  if (!p || !p.code) return { success: false, expired: true, message: 'This sign-in expired. Please sign in again.' };
  if (Date.now() > Number(p.exp)) return { success: false, message: 'The code expired. Tap “Send a new code”.' };
  if (String(code || '').replace(/\D/g, '') !== p.code) {
    p.tries = Number(p.tries || 0) + 1;
    if (p.tries >= LC_MAX_TRIES_) {
      lcDrop_(pid);
      if (typeof auditLog_ === 'function') auditLog_(p.user, 'Sign-in code failed', { details: LC_MAX_TRIES_ + ' wrong codes' });
      try { tgApi_('sendMessage', { chat_id: p.chat, text: '⚠️ حد كتب كود الدخول غلط ' + LC_MAX_TRIES_ + ' مرات على حسابك في EDEN BI. لو مش إنت، غيّر الباسورد.' }); } catch (e) {}
      return { success: false, expired: true, message: 'Too many wrong codes. Please sign in again.' };
    }
    lcPut_(pid, p);
    return { success: false, message: 'Wrong code. ' + (LC_MAX_TRIES_ - p.tries) + ' tries left.' };
  }
  lcDrop_(pid);
  const res = loginIssue_(p.user, 'Signed in (Telegram code)');
  if (remember) res.device = lcRemember_(p.user.username);
  return res;
}

/* ---------- admin (System › Users) ---------- */
function getLoginCodeSettings(authToken) {
  assertUserAdmin_(authToken);
  return { mode: lcMode_(), bot: typeof tgToken_ === 'function' && tgToken_() ? lcBotName_() : '' };
}
function setLoginCodeMode(authToken, mode) {
  const admin = assertUserAdmin_(authToken);
  mode = String(mode || '').toLowerCase();
  if (['off', 'admins', 'all'].indexOf(mode) < 0) throw new Error('Unknown setting.');
  PropertiesService.getScriptProperties().setProperty('LOGIN_2FA', mode);
  auditLog_(admin, 'Sign-in code setting', { details: mode });
  return { mode: mode };
}
function unlinkUserTelegram(authToken, username) {
  const admin = assertUserAdmin_(authToken);
  username = String(username || '').trim().toLowerCase();
  lcSetChat_(username, '');
  lcForgetDevices_(username);
  auditLog_(admin, 'Telegram unlinked for sign-in', { details: username });
  return { ok: true };
}
