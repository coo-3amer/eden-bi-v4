/************************************************
 * EDEN BI V4 — AnalyticsReport.js
 *
 * Weekly analytics report on Telegram, every Sunday (run from dailyJobs_):
 * last 7 days vs the 7 days before — sales, deals, cancellations, top
 * salespeople, cancellations per salesperson, teams, month vs target and
 * deals that are stuck. Turn off with Script Property WEEKLY_AN_OFF=1
 * (or from Analytics › Weekly report).
 ************************************************/

function weeklyAnalyticsJob_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('WEEKLY_AN_OFF') === '1') return false;
  const now = dailyLocal_(new Date());
  if (now.getDay() !== 0) return false;                       // Sunday only
  const key = Utilities.formatDate(new Date(), DAILY_TZ_, 'yyyy-MM-dd');
  if (props.getProperty('WEEKLY_AN_DAY') === key) return false;
  props.setProperty('WEEKLY_AN_DAY', key);
  return weeklyAnalyticsReport_();
}

function weeklyAnDate_(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v || '').trim(); if (!s) return null;
  const m = s.match(/^(\d{1,2})[-\/ ]([A-Za-z]{3})[A-Za-z]*[-\/ ](\d{4})$/);
  if (m) { const d = new Date(m[2] + ' ' + m[1] + ', ' + m[3]); return isNaN(d) ? null : d; }
  const d = new Date(s); return isNaN(d) ? null : d;
}
function weeklyAnChg_(a, b) {
  if (!b && !a) return '—';
  if (!b) return 'جديد';
  const d = Math.round((a - b) / Math.abs(b) * 100);
  return d === 0 ? '±0%' : (d > 0 ? '▲ ' : '▼ ') + Math.abs(d) + '%';
}
function weeklyAnMoney_(o) {
  return [o.EGP ? dailyM_(o.EGP) + ' EGP' : '', o.USD ? dailyM_(o.USD) + ' USD' : ''].filter(Boolean).join(' · ') || '0';
}

function weeklyAnalyticsReport_() {
  const now = dailyLocal_(new Date());
  const t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const DAY = 864e5;
  const from = new Date(t0.getTime() - 7 * DAY), pFrom = new Date(t0.getTime() - 14 * DAY);
  const ymd = d => d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
  const a = ymd(from), b = ymd(new Date(t0.getTime() - DAY)), pa = ymd(pFrom), pb = ymd(new Date(from.getTime() - DAY));
  const rows = dailyRows_();
  const inWk = (r, x, y) => { const k = dailyYmd_(r); return k >= x && k <= y; };
  const cur = rows.filter(r => inWk(r, a, b)), prev = rows.filter(r => inWk(r, pa, pb));
  const sum = list => { const o = { EGP: 0, USD: 0, n: 0 }; list.forEach(r => { if (dailyCancelled_(r)) return; o[dailyCur_(r)] += Number(r.amount) || 0; o.n++; }); return o; };
  const A = sum(cur), B = sum(prev);
  // cancellations that happened this week (by cancel date, else by deal date)
  const cancWk = rows.filter(r => {
    if (!dailyCancelled_(r)) return false;
    const d = weeklyAnDate_(r.cancelDate);
    return d ? (ymd(d) >= a && ymd(d) <= b) : inWk(r, a, b);
  });
  const credits = r => {
    const s = String(r.salesName || '').trim() || 'Not assigned';
    return dailyShare_(r) ? [[s, .5], [String(r.sharedWith).trim(), .5]] : [[s, 1]];
  };
  const L = [];
  const lbl = d => d.getDate() + ' ' + AR_MONTHS_[d.getMonth()];
  L.push('📈 <b>التقرير الأسبوعي للتحليلات</b>');
  L.push(lbl(from) + ' → ' + lbl(new Date(t0.getTime() - DAY)) + ' (مقارنة بالأسبوع اللي قبله)');
  L.push('');
  L.push('💰 <b>المبيعات:</b> ' + weeklyAnMoney_(A));
  L.push('   EGP ' + weeklyAnChg_(A.EGP, B.EGP) + ' · USD ' + weeklyAnChg_(A.USD, B.USD));
  L.push('🤝 <b>عدد الديلات:</b> ' + A.n + ' (' + weeklyAnChg_(A.n, B.n) + ' — كانت ' + B.n + ')');
  const cClients = {}; cancWk.forEach(r => { cClients[String(r.clientName || '').split(' - ')[0].trim().toLowerCase()] = 1; });
  L.push('❌ <b>كانسل الأسبوع:</b> ' + cancWk.length + ' ديل · ' + Object.keys(cClients).length + ' عميل · ' + weeklyAnMoney_(cancWk.reduce((o, r) => { o[dailyCur_(r)] += Number(r.amount) || 0; return o; }, { EGP: 0, USD: 0 })));

  // month to date vs company target
  const Y = now.getFullYear(), M = now.getMonth() + 1;
  const mtd = rows.filter(r => !dailyCancelled_(r) && Number(r.year) === Y && Number(r.month) === M && dailyCur_(r) === 'EGP').reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const tg = dailyCompanyTarget_(Y, M);
  const dim = new Date(Y, M, 0).getDate(), dd = Math.max(1, now.getDate() - 1);
  if (tg) {
    const pace = mtd / dd * dim;
    L.push('');
    L.push('🎯 <b>الشهر لحد دلوقتي:</b> ' + dailyM_(mtd) + ' من ' + dailyM_(tg) + ' EGP (' + Math.round(mtd / tg * 100) + '%)');
    L.push('   بالمعدل الحالي هنقفل الشهر على ' + dailyM_(pace) + ' (' + Math.round(pace / tg * 100) + '% من التارجت)');
  }

  // top salespeople this week
  const P = {};
  cur.forEach(r => credits(r).forEach(([s, w]) => {
    const o = P[s] = P[s] || { s: s, n: 0, EGP: 0, USD: 0, c: 0, all: 0, cl: {} };
    o.all += w;
    if (dailyCancelled_(r)) return;
    o.n += w; o[dailyCur_(r)] += (Number(r.amount) || 0) * w;
  }));
  cancWk.forEach(r => credits(r).forEach(([s, w]) => {
    const o = P[s] = P[s] || { s: s, n: 0, EGP: 0, USD: 0, c: 0, all: 0, cl: {} };
    o.c += w; o.cl[String(r.clientName || '').split(' - ')[0].trim().toLowerCase()] = 1;
  }));
  const top = Object.values(P).filter(o => o.n > 0).sort((x, y) => (y.EGP + y.USD * 50) - (x.EGP + x.USD * 50)).slice(0, 5);
  if (top.length) {
    L.push('');
    L.push('🏆 <b>أعلى السيلز الأسبوع ده</b>');
    top.forEach((o, i) => L.push((i + 1) + '. ' + dailyE_(o.s) + ' — ' + (Math.round(o.n * 10) / 10) + ' ديل · ' + weeklyAnMoney_(o)));
  }
  const canc = Object.values(P).filter(o => o.c > 0).sort((x, y) => y.c - x.c).slice(0, 6);
  if (canc.length) {
    L.push('');
    L.push('⚠️ <b>الكانسل لكل سيلز</b>');
    canc.forEach(o => L.push('• ' + dailyE_(o.s) + ' — ' + (Math.round(o.c * 10) / 10) + ' ديل · ' + Object.keys(o.cl).length + ' عميل'));
  }
  // teams
  const T = {};
  cur.forEach(r => { if (dailyCancelled_(r)) return; const g = String(r.salesManager || '').trim() || 'بدون مدير'; const o = T[g] = T[g] || { g: g, n: 0, EGP: 0, USD: 0 }; o.n++; o[dailyCur_(r)] += Number(r.amount) || 0; });
  const teams = Object.values(T).sort((x, y) => (y.EGP + y.USD * 50) - (x.EGP + x.USD * 50)).slice(0, 6);
  if (teams.length) {
    L.push('');
    L.push('👥 <b>الفرق</b>');
    teams.forEach(o => L.push('• ' + dailyE_(o.g) + ' — ' + o.n + ' ديل · ' + weeklyAnMoney_(o)));
  }
  // stuck deals
  const age = r => { const y = Number(r.year), m = Number(r.month), d = Number(r.day) || 1; return y && m ? Math.floor((t0 - new Date(y, m - 1, d)) / DAY) : 0; };
  const st = r => String(r.status || '').trim().toLowerCase();
  const res7 = rows.filter(r => !dailyCancelled_(r) && st(r) === 'reservation' && age(r) > 7).length;
  const dp30 = rows.filter(r => !dailyCancelled_(r) && /part of dp|total dp/.test(st(r)) && age(r) > 30).length;
  const openDp = rows.filter(r => !dailyCancelled_(r) && (Number(r.actualDP) || 0) > (Number(r.dpPaid) || 0) + 1);
  L.push('');
  L.push('⏳ <b>محتاج متابعة</b>');
  L.push('• حجوزات عدّى عليها أكتر من 7 أيام: ' + res7);
  L.push('• دفعة مقدمة من غير توقيع أكتر من 30 يوم: ' + dp30);
  L.push('• ديلات لسه فيها مقدم مش مدفوع: ' + openDp.length);
  L.push('');
  L.push('التفاصيل كاملة في EDEN BI › Analytics');
  return dailySend_(L.join('\n'));
}

/* ---------- Admin (Analytics page) ---------- */
function getWeeklyAnalyticsInfo(authToken) {
  targetsRequireAdmin_(authToken);
  return { on: PropertiesService.getScriptProperties().getProperty('WEEKLY_AN_OFF') !== '1' };
}
function setWeeklyAnalytics(authToken, on) {
  const user = targetsRequireAdmin_(authToken);
  const props = PropertiesService.getScriptProperties();
  if (on) props.deleteProperty('WEEKLY_AN_OFF'); else props.setProperty('WEEKLY_AN_OFF', '1');
  auditLog_(user, 'Weekly analytics report', { details: on ? 'on' : 'off' });
  return { on: !!on };
}
function sendWeeklyAnalyticsNow(authToken) {
  const user = targetsRequireAdmin_(authToken);
  const ok = weeklyAnalyticsReport_();
  auditLog_(user, 'Weekly analytics report sent', { details: ok ? 'sent' : 'not sent' });
  return { ok: !!ok };
}
