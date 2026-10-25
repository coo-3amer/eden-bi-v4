/************************************************
 * EDEN BI V4 — DailyJobs.js
 *
 * Once a day (from 9:00 Cairo time, run by the 5-minute warm-up):
 *  1) Morning brief for managers on Telegram — yesterday's deals, the
 *     month against the company target, what is needed per day, and
 *     reservations ending today / tomorrow, installments due today.
 *  2) Collection reminders — installments due in 3 days / today, and
 *     overdue ones on day 1, 7, 14, 30, 60, 90… so nobody is spammed.
 * Messages go to TELEGRAM_CHAT_ID; reminders go to
 * TELEGRAM_COLLECTIONS_CHAT_ID instead when that property is set.
 * Turn either off with Script Properties BRIEF_OFF=1 / REMIND_OFF=1.
 ************************************************/

const DAILY_HOUR_ = 9;
const DAILY_TZ_ = 'Africa/Cairo';
const AR_MONTHS_ = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
const AR_DAYS_ = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

function dailyJobs_(force) {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const day = Utilities.formatDate(now, DAILY_TZ_, 'yyyy-MM-dd');
  const hour = Number(Utilities.formatDate(now, DAILY_TZ_, 'H'));
  if (!force && (hour < DAILY_HOUR_ || props.getProperty('DAILY_JOBS_DAY') === day)) return;
  props.setProperty('DAILY_JOBS_DAY', day);
  if (props.getProperty('BRIEF_OFF') !== '1') { try { morningBrief_(); } catch (e) { console.warn('Morning brief: ' + e); } }
  if (props.getProperty('REMIND_OFF') !== '1') { try { collectionReminders_(); } catch (e) { console.warn('Collection reminders: ' + e); } }
  try { if (typeof boardPackMonthly_ === 'function') boardPackMonthly_(); } catch (e) { console.warn('Board pack: ' + e); }
}

/* ---------- helpers ---------- */
function dailySend_(html, chatProp) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('TELEGRAM_OFF') === '1') return false;
  const chat = String((chatProp && props.getProperty(chatProp)) || props.getProperty('TELEGRAM_CHAT_ID') || '').trim();
  if (!chat || typeof tgApi_ !== 'function') return false;
  // Telegram limit ~4096 chars: split on line breaks.
  const parts = [], lines = String(html).split('\n');
  let cur = '';
  lines.forEach(l => { if ((cur + '\n' + l).length > 3800) { parts.push(cur); cur = l; } else cur = cur ? cur + '\n' + l : l; });
  if (cur) parts.push(cur);
  let ok = true;
  parts.forEach(p => { const r = tgApi_('sendMessage', { chat_id: chat, text: p, parse_mode: 'HTML', disable_web_page_preview: true }); if (!r || !r.ok) ok = false; });
  return ok;
}
function dailyE_(v) { return typeof tgEsc_ === 'function' ? tgEsc_(v) : String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function dailyN_(n) { return Math.round(Number(n) || 0).toLocaleString('en-US'); }
function dailyM_(n) {   // 12.5M / 850K
  const v = Number(n) || 0, a = Math.abs(v);
  if (a >= 1e6) return String(Math.round(v / 1e4) / 100) + 'M';
  if (a >= 1e3) return Math.round(v / 1e3) + 'K';
  return String(Math.round(v));
}
function dailyLocal_(d) { return new Date(Utilities.formatDate(d, DAILY_TZ_, "yyyy-MM-dd'T'HH:mm:ss")); }
function dailyDateLabel_(d) { const x = dailyLocal_(d); return AR_DAYS_[x.getDay()] + ' ' + x.getDate() + ' ' + AR_MONTHS_[x.getMonth()]; }
function dailyRows_() {
  let data = typeof dashCacheRead_ === 'function' ? dashCacheRead_('ALL', 6 * 3600 * 1000) : null;
  if (!data) data = buildDashboardData_({ role: 'Admin' });
  return (data && data.rows) || [];
}
function dailyCancelled_(r) { return typeof cancelledStatus_ === 'function' ? cancelledStatus_(r.status) : /^cancel/i.test(String(r.status || '')); }
function dailyTrue_(v) { if (v === true || v === 1) return true; const s = String(v == null ? '' : v).trim().toLowerCase(); return s === 'true' || s === '1' || s === 'yes' || s === '✓'; }
function dailyCur_(r) { const c = String(r.currency || '').toUpperCase(); return c.indexOf('USD') >= 0 || c === '$' ? 'USD' : 'EGP'; }
function dailyShare_(r) { return /share/i.test(String(r.dealStatus || '')) && String(r.sharedWith || '').trim() && String(r.sharedWith).trim().toLowerCase() !== String(r.salesName || '').trim().toLowerCase(); }
function dailyYmd_(r) { const y = Number(r.year), m = Number(r.month), d = Number(r.day); return y && m ? y * 10000 + m * 100 + (d || 1) : 0; }

/* Company target for a month: monthly, else quarter ÷ 3, else year ÷ 12 (EGP). */
function dailyCompanyTarget_(year, month) {
  let rows = [];
  try { rows = getTargetsRows_().rows || []; } catch (e) { return 0; }
  const mine = rows.filter(t => String(t.scope).toLowerCase() === 'company' && Number(t.year) === year);
  const mName = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'][month - 1];
  const sum = list => list.reduce((a, t) => a + (Number(t.targetEGP) || 0), 0);
  const mo = sum(mine.filter(t => String(t.period).toLowerCase() === 'monthly' && String(t.month).trim().toLowerCase() === mName));
  if (mo) return mo;
  const q = sum(mine.filter(t => String(t.period).toLowerCase() === 'quarterly' && String(t.quarter).toUpperCase() === 'Q' + Math.ceil(month / 3)));
  if (q) return q / 3;
  const y = sum(mine.filter(t => String(t.period).toLowerCase() === 'yearly'));
  return y ? y / 12 : 0;
}

/* ---------- 1) Morning brief ---------- */
function morningBrief_() {
  const now = dailyLocal_(new Date());
  const yest = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const yKey = yest.getFullYear() * 10000 + (yest.getMonth() + 1) * 100 + yest.getDate();
  const Y = now.getFullYear(), M = now.getMonth() + 1;
  const rows = dailyRows_();
  const lines = [];
  lines.push('☀️ <b>صباح الخير — ملخص EDEN</b>');
  lines.push(dailyDateLabel_(new Date()));
  lines.push('');

  // Yesterday
  const yd = rows.filter(r => dailyYmd_(r) === yKey && Number(r.amount) > 0);
  const yOk = yd.filter(r => !dailyCancelled_(r));
  const tot = (list, cur) => list.filter(r => dailyCur_(r) === cur).reduce((a, r) => a + (Number(r.amount) || 0), 0);
  lines.push('📅 <b>امبارح (' + dailyDateLabel_(yest) + ')</b>');
  if (!yOk.length) lines.push('مفيش ديلات جديدة.');
  else {
    const usd = tot(yOk, 'USD');
    lines.push(yOk.length + ' ديل · <b>' + dailyM_(tot(yOk, 'EGP')) + ' EGP</b>' + (usd ? ' · ' + dailyM_(usd) + ' USD' : ''));
    yOk.slice(0, 12).forEach(r => lines.push('• ' + dailyE_(r.unitCode) + ' — ' + dailyE_(r.salesName) + (dailyShare_(r) ? ' + ' + dailyE_(r.sharedWith) : '') + ' · ' + dailyM_(r.amount) + ' ' + dailyCur_(r) + ' · ' + dailyE_(r.status)));
    if (yOk.length > 12) lines.push('… و' + (yOk.length - 12) + ' ديل تاني');
  }
  lines.push('');

  // Month to date (EGP)
  const mRows = rows.filter(r => Number(r.year) === Y && Number(r.month) === M && !dailyCancelled_(r) && Number(r.amount) > 0 && dailyCur_(r) === 'EGP');
  const booked = mRows.reduce((a, r) => a + (Number(r.amount) || 0), 0);
  const signed = mRows.filter(r => dailyTrue_(r.confClientSigned)).reduce((a, r) => a + (Number(r.amount) || 0), 0);
  const target = dailyCompanyTarget_(Y, M);
  const daysInMonth = new Date(Y, M, 0).getDate(), daysLeft = daysInMonth - now.getDate() + 1;
  lines.push('📈 <b>شهر ' + AR_MONTHS_[M - 1] + ' لحد دلوقتي (EGP)</b>');
  lines.push('المبيعات: <b>' + dailyM_(booked) + '</b> (' + mRows.length + ' ديل) · العميل وقّع: <b>' + dailyM_(signed) + '</b>');
  if (target) {
    const pct = Math.round(signed / target * 1000) / 10, left = Math.max(0, target - signed);
    const pace = Math.round((now.getDate() - 1) / daysInMonth * 100);
    lines.push('التارجت: ' + dailyM_(target) + ' · المحقق <b>' + pct + '%</b> (المفروض نكون ' + pace + '%)' + (pct >= pace ? ' ✅' : ' ⚠️'));
    if (left > 0) lines.push('الباقي ' + dailyM_(left) + ' في ' + daysLeft + ' يوم → <b>' + dailyM_(left / Math.max(1, daysLeft)) + ' يومياً</b>');
    else lines.push('🎉 التارجت اتحقق!');
  } else lines.push('مفيش تارجت للشركة متسجل للشهر ده.');
  // Top 3 salespeople this month (share deals split 50/50)
  const by = {};
  mRows.forEach(r => {
    const a = Number(r.amount) || 0;
    if (dailyShare_(r)) { by[r.salesName] = (by[r.salesName] || 0) + a / 2; by[r.sharedWith] = (by[r.sharedWith] || 0) + a / 2; }
    else if (r.salesName) by[r.salesName] = (by[r.salesName] || 0) + a;
  });
  const top = Object.keys(by).filter(n => n && n.toLowerCase() !== 'not assigned').sort((a, b) => by[b] - by[a]).slice(0, 3);
  if (top.length) lines.push('🏆 ' + top.map((n, i) => ['🥇', '🥈', '🥉'][i] + ' ' + dailyE_(n) + ' ' + dailyM_(by[n])).join('  '));
  lines.push('');

  // Reservations ending today / tomorrow
  try {
    const hl = resHoldList_();
    if (!hl.auto) {
      // manual only: list reservations that are 7+ days old so someone decides on them
      const lateR = (hl.rows || []).filter(h => h.days != null && h.days >= (hl.holdDays || 7)).sort((x, y) => y.days - x.days);
      if (lateR.length) {
        lines.push('⏳ <b>حجوزات عدّى عليها ' + (hl.holdDays || 7) + ' أيام أو أكتر (' + lateR.length + ') — محتاجة قرار يدوي</b>');
        lateR.slice(0, 15).forEach(h => lines.push('• ' + dailyE_(h.unit) + ' — ' + dailyE_(h.client) + ' · ' + dailyE_(h.sales) + ' · ' + h.days + ' يوم'));
        lines.push('');
      }
    }
    const holds = (hl.rows || []).filter(h => h.endsAt);
    const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 2).getTime();
    const soon = holds.filter(h => new Date(h.endsAt).getTime() < end).sort((a, b) => new Date(a.endsAt) - new Date(b.endsAt));
    if (soon.length) {
      lines.push('⏳ <b>حجوزات بتنتهي النهارده وبكرة (' + soon.length + ')</b>');
      soon.slice(0, 15).forEach(h => lines.push('• ' + dailyE_(h.unit) + ' — ' + dailyE_(h.client) + ' · ' + dailyE_(h.sales) + ' · ' + Utilities.formatDate(new Date(h.endsAt), DAILY_TZ_, 'd/M HH:mm')));
      lines.push('');
    }
  } catch (e) { console.warn('Brief holds: ' + e); }

  // Installments due today + overdue total
  try {
    const t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    const inst = dailyInst_();
    const dueToday = inst.filter(x => x.t === t0), late = inst.filter(x => x.t < t0);
    if (dueToday.length || late.length) {
      lines.push('💳 <b>التحصيل</b>');
      if (dueToday.length) lines.push('مستحق النهارده: ' + dueToday.length + ' قسط · ' + dailyTotals_(dueToday));
      if (late.length) lines.push('إجمالي المتأخر: ' + late.length + ' قسط · ' + dailyTotals_(late));
    }
  } catch (e) { console.warn('Brief installments: ' + e); }

  return dailySend_(lines.join('\n'));
}

/* ---------- 2) Collection reminders ---------- */
function dailyInst_() {
  const ov = instOverview_();
  return (ov.rows || []).map(x => {
    const d = instDateIn_(x.due);
    return Object.assign({}, x, { left: Math.max(0, (Number(x.amount) || 0) - (Number(x.paid) || 0)), t: d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() : 0 });
  }).filter(x => x.left > 1 && x.t);
}
function dailyTotals_(list) {
  const e = list.filter(x => String(x.currency || 'EGP').toUpperCase() !== 'USD').reduce((a, x) => a + x.left, 0);
  const u = list.filter(x => String(x.currency || '').toUpperCase() === 'USD').reduce((a, x) => a + x.left, 0);
  return [e ? dailyN_(e) + ' EGP' : '', u ? dailyN_(u) + ' USD' : ''].filter(Boolean).join(' · ') || '0';
}
function collectionReminders_() {
  const now = dailyLocal_(new Date());
  const t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime(), DAY = 864e5;
  const inst = dailyInst_();
  if (!inst.length) return false;
  const nudge = [1, 7, 14, 30, 60, 90];
  const in3 = inst.filter(x => x.t === t0 + 3 * DAY), today = inst.filter(x => x.t === t0);
  const late = inst.filter(x => x.t < t0).map(x => Object.assign(x, { days: Math.round((t0 - x.t) / DAY) }));
  const lateNudge = late.filter(x => nudge.indexOf(x.days) >= 0 || (x.days > 90 && x.days % 30 === 0)).sort((a, b) => b.days - a.days);
  if (!in3.length && !today.length && !lateNudge.length) return false;
  const line = x => '• ' + dailyE_(x.unitCode) + ' — ' + dailyE_(x.client) + ' · ' + dailyE_(x.name) + ' · <b>' + dailyN_(x.left) + ' ' + dailyE_(x.currency || 'EGP') + '</b>';
  const L = ['💳 <b>تذكير التحصيل</b> · ' + dailyDateLabel_(new Date()), ''];
  if (in3.length) { L.push('⏰ <b>مستحق بعد 3 أيام (' + in3.length + ')</b>'); in3.slice(0, 25).forEach(x => L.push(line(x))); L.push(''); }
  if (today.length) { L.push('📌 <b>مستحق النهارده (' + today.length + ')</b>'); today.slice(0, 25).forEach(x => L.push(line(x))); L.push(''); }
  if (lateNudge.length) { L.push('🔴 <b>متأخر (' + lateNudge.length + ')</b>'); lateNudge.slice(0, 30).forEach(x => L.push(line(x) + ' · ' + x.days + ' يوم')); L.push(''); }
  L.push('إجمالي المتأخر كله: ' + late.length + ' قسط · ' + dailyTotals_(late));
  return dailySend_(L.join('\n'), 'TELEGRAM_COLLECTIONS_CHAT_ID');
}

/* ---------- Admin: send now (buttons in the dashboard) ---------- */
function sendMorningBriefNow(authToken) {
  const user = targetsRequireAdmin_(authToken);
  const ok = morningBrief_();
  auditLog_(user, 'Morning brief sent', { details: ok ? 'sent' : 'not sent' });
  return { ok: !!ok };
}
function sendCollectionRemindersNow(authToken) {
  const user = targetsRequireAdmin_(authToken);
  const ok = collectionReminders_();
  auditLog_(user, 'Collection reminders sent', { details: ok ? 'sent' : 'nothing due' });
  return { ok: !!ok };
}
