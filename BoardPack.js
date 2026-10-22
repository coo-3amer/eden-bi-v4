/************************************************
 * EDEN BI V4 — BoardPack.js
 *
 * Monthly management report for the month that just ended: sales vs the
 * month before and the same month last year, target, projects, top
 * salespeople and sources, cancellations, collections (received, overdue,
 * expected), and internal commissions if the month is closed.
 * Sent as a PDF to Telegram on the 1st of each month (from the daily job),
 * and openable any time from the Targets page (admins).
 * Set BOARD_PACK_OFF=1 in Script Properties to stop the monthly send.
 ************************************************/

function bpMonthInfo_(ym) {
  let y, m;
  if (ym && /^\d{4}-\d{2}$/.test(ym)) { y = +ym.slice(0, 4); m = +ym.slice(5); }
  else { const now = dailyLocal_(new Date()); y = now.getFullYear(); m = now.getMonth(); if (m === 0) { m = 12; y--; } }   // last month
  const key = y + '-' + String(m).padStart(2, '0');
  const prev = m === 1 ? { y: y - 1, m: 12 } : { y: y, m: m - 1 };
  return { y: y, m: m, key: key, prev: prev, ly: { y: y - 1, m: m },
    label: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][m - 1] + ' ' + y };
}

function boardPackHtml_(ym) {
  const M = bpMonthInfo_(ym), rows = dailyRows_();
  const inMonth = (r, p) => Number(r.year) === p.y && Number(r.month) === p.m && Number(r.amount) > 0;
  const ok = r => !dailyCancelled_(r);
  const sum = (list, cur) => list.filter(r => dailyCur_(r) === cur).reduce((a, r) => a + (Number(r.amount) || 0), 0);
  const cur = rows.filter(r => inMonth(r, M)), act = cur.filter(ok), canc = cur.filter(r => !ok(r));
  const prev = rows.filter(r => inMonth(r, M.prev) && ok(r)), ly = rows.filter(r => inMonth(r, M.ly) && ok(r));
  const egp = sum(act, 'EGP'), usd = sum(act, 'USD'), signed = act.filter(r => dailyTrue_(r.confClientSigned) && dailyCur_(r) === 'EGP').reduce((a, r) => a + (Number(r.amount) || 0), 0);
  const target = dailyCompanyTarget_(M.y, M.m);
  const ch = (a, b) => b ? Math.round((a - b) / b * 1000) / 10 : null;
  const chTxt = v => v == null ? '—' : (v >= 0 ? '▲ ' : '▼ ') + Math.abs(v) + '%';
  const E = dailyE_, N = dailyN_, Mn = dailyM_;
  const group = (list, keyFn) => {
    const g = {};
    list.forEach(r => {
      const a = Number(r.amount) || 0, c = dailyCur_(r);
      [].concat(keyFn(r)).forEach(([k, share]) => { if (!k) return; const o = g[k] || (g[k] = { k: k, n: 0, EGP: 0, USD: 0 }); o.n += share; o[c] += a * share; });
    });
    return Object.values(g).sort((a, b) => (b.EGP + b.USD * 50) - (a.EGP + a.USD * 50));
  };
  const byProject = group(act, r => [[r.project || '—', 1]]);
  const bySales = group(act, r => dailyShare_(r) ? [[r.salesName, .5], [r.sharedWith, .5]] : [[r.salesName, 1]]).filter(x => String(x.k).toLowerCase() !== 'not assigned').slice(0, 8);
  const bySource = group(act, r => [[r.mainSource || '—', 1]]).slice(0, 8);
  // collections
  let recvE = 0, recvU = 0, lateE = 0, next3 = 0;
  try {
    const ov = instOverview_(), rec = ov.received || {};
    recvE = rec[M.key + '|EGP'] || 0; recvU = rec[M.key + '|USD'] || 0;
    const now = dailyLocal_(new Date()), t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime(), t3 = new Date(now.getFullYear(), now.getMonth() + 3, now.getDate()).getTime();
    dailyInst_().forEach(x => { if (String(x.currency || 'EGP').toUpperCase() === 'USD') return; if (x.t < t0) lateE += x.left; else if (x.t < t3) next3 += x.left; });
  } catch (e) { console.warn('Board pack collections: ' + e); }
  // commissions (closed month)
  let comm = null;
  try {
    const cr = commCloseRows_().filter(r => String(r[0]) === M.key);
    if (cr.length) comm = { people: cr.length, total: cr.reduce((a, r) => a + (Number(r[7]) || 0), 0), paid: cr.filter(r => r[12] === true || String(r[12]).toUpperCase() === 'TRUE').reduce((a, r) => a + (Number(r[7]) || 0), 0) };
  } catch (e) {}
  const tbl = (head, list, row) => '<table><thead><tr>' + head.map(h => '<th' + (/^(Deals|EGP|USD)$/.test(h) ? ' class="n"' : '') + '>' + h + '</th>').join('') + '</tr></thead><tbody>' + list.map(row).join('') + '</tbody></table>';
  const kp = (l, v, s, cls) => '<td class="kp' + (cls ? ' ' + cls : '') + '"><span>' + l + '</span><b>' + v + '</b>' + (s ? '<small>' + s + '</small>' : '') + '</td>';
  const pct = target ? Math.round(signed / target * 1000) / 10 : null;
  return '<!doctype html><html><head><meta charset="utf-8"><title>EDEN — Monthly report ' + E(M.label) + '</title><style>' +
    '@page{size:A4;margin:12mm}body{font-family:Arial,Helvetica,sans-serif;color:#1d1a16;font-size:11px;margin:0}' +
    '.hd{border-bottom:2px solid #b08a42;padding-bottom:8px;margin-bottom:12px}.brand{font-family:Georgia,serif;font-size:24px;letter-spacing:4px;font-weight:bold}.brand small{display:block;font-family:Arial;font-size:8px;letter-spacing:5px;color:#8f6d2c}' +
    'h1{font-size:18px;margin:6px 0 0}h1 small{font-size:11px;color:#83796b;font-weight:normal}' +
    'h2{font-family:Georgia,serif;font-size:14px;margin:16px 0 6px;color:#3a3227}' +
    'table{width:100%;border-collapse:collapse}th{background:#f3ece0;text-align:left;font-size:8.5px;letter-spacing:.5px;text-transform:uppercase;color:#6f675b;padding:5px 6px}td{padding:5px 6px;border-bottom:1px solid #eee6d8}.n{text-align:right}' +
    'table.k{border-collapse:separate;border-spacing:6px;margin:0 -6px}td.kp{border:1px solid #e6dcc8;background:#fbf8f2;border-radius:8px;padding:8px 10px;vertical-align:top;width:25%}td.kp span{display:block;font-size:8.5px;font-weight:bold;letter-spacing:.5px;text-transform:uppercase;color:#8a7f70}td.kp b{display:block;font-size:15px;margin-top:2px}td.kp small{display:block;color:#83796b;margin-top:2px}td.kp.g{background:#f6eedc;border-color:#d9c08c}td.kp.r b{color:#b3261e}' +
    '.ft{margin-top:14px;font-size:8.5px;color:#8a7f70}</style></head><body>' +
    '<div class="hd"><div class="brand">EDEN<small>DEVELOPMENT</small></div><h1>Monthly management report — ' + E(M.label) + ' <small>· generated ' + Utilities.formatDate(new Date(), DAILY_TZ_, 'd MMM yyyy HH:mm') + '</small></h1></div>' +
    '<h2>Sales</h2><table class="k"><tr>' +
      kp('Sales EGP', Mn(egp), act.length + ' deals · ' + chTxt(ch(egp, sum(prev, 'EGP'))) + ' vs last month', 'g') +
      kp('Client signed EGP', Mn(signed), target ? pct + '% of target ' + Mn(target) : 'no company target') +
      kp('Same month last year', Mn(sum(ly, 'EGP')), chTxt(ch(egp, sum(ly, 'EGP')))) +
      kp('Sales USD', Mn(usd), act.filter(r => dailyCur_(r) === 'USD').length + ' deals') +
    '</tr><tr>' +
      kp('Cancelled', canc.length + ' deals', Mn(sum(canc, 'EGP')) + ' EGP of this month’s reservations', canc.length ? 'r' : '') +
      kp('Received (installments)', Mn(recvE) + ' EGP', recvU ? Mn(recvU) + ' USD' : '') +
      kp('Overdue now', Mn(lateE) + ' EGP', 'from payment schedules', lateE ? 'r' : '') +
      kp('Expected next 3 months', Mn(next3) + ' EGP', 'still unpaid') +
    '</tr></table>' +
    '<h2>By project</h2>' + tbl(['Project', 'Deals', 'EGP', 'USD'], byProject, x => '<tr><td>' + E(x.k) + '</td><td class="n">' + Math.round(x.n * 10) / 10 + '</td><td class="n">' + N(x.EGP) + '</td><td class="n">' + N(x.USD) + '</td></tr>') +
    '<h2>Top salespeople</h2>' + tbl(['#', 'Name', 'Deals', 'EGP', 'USD'], bySales, (x, i) => '<tr><td>' + (i + 1) + '</td><td>' + E(x.k) + '</td><td class="n">' + Math.round(x.n * 10) / 10 + '</td><td class="n">' + N(x.EGP) + '</td><td class="n">' + N(x.USD) + '</td></tr>') +
    '<h2>Top sources</h2>' + tbl(['Main source', 'Deals', 'EGP', 'USD'], bySource, x => '<tr><td>' + E(x.k) + '</td><td class="n">' + x.n + '</td><td class="n">' + N(x.EGP) + '</td><td class="n">' + N(x.USD) + '</td></tr>') +
    (comm ? '<h2>Internal commissions (closed)</h2><table class="k"><tr>' + kp('Total', N(comm.total) + ' EGP', comm.people + ' people', 'g') + kp('Paid', N(comm.paid) + ' EGP', '') + kp('Still to pay', N(comm.total - comm.paid) + ' EGP', '') + '<td></td></tr></table>' : '') +
    '<div class="ft">Deals by reservation date, cancelled deals excluded (except the Cancelled box). Shared deals count half for each salesperson. Target achievement uses “Client signed”. EDEN BI · confidential.</div>' +
    '</body></html>';
}

function boardPackSend_(ym) {
  const props = PropertiesService.getScriptProperties();
  const chat = String(props.getProperty('TELEGRAM_CHAT_ID') || '').trim(), token = typeof tgToken_ === 'function' ? tgToken_() : '';
  if (!chat || !token || props.getProperty('TELEGRAM_OFF') === '1') return false;
  const M = bpMonthInfo_(ym);
  const pdf = Utilities.newBlob(boardPackHtml_(ym), 'text/html', 'report.html').getAs('application/pdf').setName('EDEN monthly report ' + M.key + '.pdf');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendDocument', {
    method: 'post', muteHttpExceptions: true,
    payload: { chat_id: chat, caption: '📊 التقرير الشهري للإدارة — ' + M.label, document: pdf }
  });
  const ok = res.getResponseCode() === 200;
  if (!ok) console.warn('Board pack send: ' + res.getContentText().slice(0, 300));
  return ok;
}

/* From the daily job: on the 1st, once per month. */
function boardPackMonthly_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('BOARD_PACK_OFF') === '1') return;
  const now = dailyLocal_(new Date());
  if (now.getDate() !== 1) return;
  const M = bpMonthInfo_();
  if (props.getProperty('BOARD_PACK_SENT') === M.key) return;
  if (boardPackSend_()) props.setProperty('BOARD_PACK_SENT', M.key);
}

/* Admin: open (html) and/or send now. */
function getBoardPack(authToken, ym, send) {
  const user = targetsRequireAdmin_(authToken);
  const html = boardPackHtml_(ym);
  let sent = false;
  if (send) { sent = boardPackSend_(ym); auditLog_(user, 'Monthly report sent', { details: bpMonthInfo_(ym).label + (sent ? '' : ' · not sent') }); }
  return { html: html, sent: sent, label: bpMonthInfo_(ym).label };
}
